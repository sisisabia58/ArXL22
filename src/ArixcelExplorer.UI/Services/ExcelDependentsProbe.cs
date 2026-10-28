using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.UI.Services;

/// <summary>
/// Reads direct dependents from the running Excel instance. The installed host
/// walks DirectDependents.Cells, which splits every arrow target into single
/// cells and never sees other sheets. Areas keep G48:G53 together, and
/// ShowDependents/NavigateArrow supplies other-sheet targets such as
/// 'Summary (Instit)'!D11:J11. Arrows are cleared and the selection restored.
/// </summary>
internal static class ExcelDependentsProbe
{
    public static IReadOnlyList<DependentEntry>? TryRead(string? originAddress)
    {
        if (string.IsNullOrWhiteSpace(originAddress)) return null;

        object excel;
        try
        {
            excel = Marshal.GetActiveObject("Excel.Application");
        }
        catch
        {
            return null;
        }

        var origin = RangeOf(excel, originAddress!);
        if (origin == null) return null;

        object? previousUpdating = null;
        object? previousEvents = null;
        object? originSheet = null;
        var restoreAddress = originAddress!;
        var arrows = new List<string>();
        var areas = new List<string>();
        try
        {
            try { previousUpdating = Get(excel, "ScreenUpdating"); } catch { /* keep going */ }
            try { previousEvents = Get(excel, "EnableEvents"); } catch { /* keep going */ }
            try { Set(excel, "ScreenUpdating", false); } catch { /* keep going */ }
            try { Set(excel, "EnableEvents", false); } catch { /* keep going */ }

            try
            {
                var active = Get(excel, "ActiveCell");
                var activeAddress = FormatRange(active);
                if (!string.IsNullOrEmpty(activeAddress)) restoreAddress = activeAddress;
            }
            catch
            {
                // The origin address is still a safe place to put the selection back.
            }

            try { originSheet = Get(origin, "Worksheet"); } catch { /* cleared below if present */ }
            try { ReadAreas(origin, areas); } catch { /* no same-sheet dependents */ }
            try
            {
                // One select so the tracer starts at the source. Selecting again
                // between arrows clears them and drops later areas such as D67:J67.
                try { Invoke(origin, "Select"); } catch { /* ShowDependents still accepts the range */ }
                if (ShowDependents(origin))
                {
                    WalkArrows(excel, origin, originAddress!, arrows);
                }
            }
            catch
            {
                // Cross-sheet arrows are optional; same-sheet areas still stand.
            }

            try
            {
                // DirectDependents can stop at column G while the arrow selection,
                // array, or neighboring formulas still cover D:J and row 67.
                var discovered = new List<string>(arrows.Count + areas.Count);
                discovered.AddRange(arrows);
                discovered.AddRange(areas);
                ExpandByFormulas(excel, originAddress!, discovered);
                areas.AddRange(discovered);
            }
            catch
            {
                // Keep whatever Areas and arrows were already collected.
            }
        }
        finally
        {
            try
            {
                if (originSheet != null) Invoke(originSheet, "ClearArrows");
            }
            catch
            {
                // A missing arrow set should not hide the dependents window.
            }

            try
            {
                var back = RangeOf(excel, restoreAddress);
                if (back != null) Invoke(back, "Select");
            }
            catch
            {
                // Selection restore is best-effort.
            }

            try
            {
                if (previousEvents != null) Set(excel, "EnableEvents", previousEvents);
            }
            catch
            {
                // Events are restored when Excel allows it.
            }

            try
            {
                if (previousUpdating != null) Set(excel, "ScreenUpdating", previousUpdating);
            }
            catch
            {
                // Leave Excel able to paint if the property rejects the old value.
            }
        }

        return DependentsTree.Combine(originAddress, arrows, areas);
    }

    private static void ReadAreas(object origin, List<string> areas)
    {
        var deps = Get(origin, "DirectDependents");
        if (deps == null) return;
        foreach (var piece in ListAddresses(deps))
        {
            if (!string.IsNullOrEmpty(piece)) areas.Add(piece);
        }
    }

    private static bool ShowDependents(object origin)
    {
        try
        {
            Invoke(origin, "ShowDependents", false);
            return true;
        }
        catch
        {
            try
            {
                Invoke(origin, "ShowDependents");
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    private static void WalkArrows(object excel, object originRange, string originAddress, List<string> arrows)
    {
        var originNorm = DependentsTree.NormalizeAddress(originAddress, originAddress);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var duplicateArrows = 0;
        var source = originRange;
        for (var arrow = 1; arrow <= 1024; arrow++)
        {
            var added = 0;
            var linkWorked = false;
            for (var link = 1; link <= 16; link++)
            {
                object? dest;
                try
                {
                    // A second call after a failure restarts the arrow list and
                    // skips the last area. Stop this link, or the walk, instead.
                    dest = Invoke(source, "NavigateArrow", false, arrow, link);
                }
                catch
                {
                    if (link == 1) return;
                    break;
                }

                if (dest == null)
                {
                    if (link == 1) return;
                    break;
                }

                linkWorked = true;
                foreach (var piece in DescribeTarget(excel, dest))
                {
                    var norm = DependentsTree.NormalizeAddress(piece, originAddress);
                    if (string.IsNullOrEmpty(norm) ||
                        TraceUtils.AddressesReferToSameRange(norm, originNorm))
                    {
                        continue;
                    }

                    if (seen.Add(norm))
                    {
                        arrows.Add(piece);
                        added++;
                    }
                }
            }

            if (!linkWorked) return;
            if (added == 0)
            {
                duplicateArrows++;
                if (duplicateArrows >= 2) return;
            }
            else
            {
                duplicateArrows = 0;
            }
        }
    }

    private static List<string> DescribeTarget(object excel, object dest)
    {
        var pieces = ListAddresses(dest);
        var extras = new List<string>();
        try
        {
            var selection = Get(excel, "Selection");
            if (selection != null) extras.AddRange(ListAddresses(selection));
        }
        catch
        {
            // Selection is only a wider view of the same arrow.
        }

        try
        {
            var currentArray = Get(dest, "CurrentArray");
            if (currentArray != null) extras.AddRange(ListAddresses(currentArray));
        }
        catch
        {
            // CurrentArray throws when the target is not an array formula.
        }

        if (pieces.Count == 0) return extras;
        var widened = new List<string>(pieces.Count);
        foreach (var piece in pieces)
        {
            var best = piece;
            foreach (var extra in extras)
            {
                best = DependentsTree.WidenSingleRow(best, extra);
            }

            widened.Add(best);
        }

        foreach (var extra in extras)
        {
            var related = false;
            foreach (var piece in pieces)
            {
                var extraNorm = DependentsTree.NormalizeAddress(extra, piece);
                var pieceNorm = DependentsTree.NormalizeAddress(piece, piece);
                if (DependentsTree.ContainsRange(extraNorm, pieceNorm) ||
                    DependentsTree.ContainsRange(pieceNorm, extraNorm) ||
                    TraceUtils.AddressesReferToSameRange(extraNorm, pieceNorm))
                {
                    related = true;
                    break;
                }
            }

            if (!related) continue;
            var covered = false;
            foreach (var piece in widened)
            {
                if (TraceUtils.AddressesReferToSameRange(piece, extra) ||
                    DependentsTree.ContainsRange(piece, DependentsTree.NormalizeAddress(extra, piece)))
                {
                    covered = true;
                    break;
                }
            }

            if (!covered) widened.Add(extra);
        }

        return widened;
    }

    private static List<string> ListAddresses(object range)
    {
        try
        {
            var areas = Get(range, "Areas");
            if (areas != null)
            {
                var count = Convert.ToInt32(Get(areas, "Count"), CultureInfo.InvariantCulture);
                if (count > 1)
                {
                    var pieces = new List<string>();
                    for (var i = 1; i <= count; i++)
                    {
                        var area = Item(areas, i);
                        if (area == null) continue;
                        var text = AddressOf(area);
                        if (!string.IsNullOrEmpty(text)) pieces.Add(text);
                    }

                    if (pieces.Count > 0) return pieces;
                }
            }
        }
        catch
        {
            // Fall through to the range address.
        }

        var single = AddressOf(range);
        if (!string.IsNullOrEmpty(single)) return SplitQualified(single);
        try
        {
            var formatted = FormatRange(range);
            if (!string.IsNullOrEmpty(formatted)) return SplitQualified(formatted);
        }
        catch
        {
            // Unreadable COM range.
        }

        return new List<string>();
    }

    private static List<string> SplitQualified(string address)
    {
        var pieces = new List<string>();
        foreach (var part in DependentsTree.SplitAreas(address))
        {
            if (!string.IsNullOrWhiteSpace(part)) pieces.Add(part.Trim());
        }

        return pieces.Count > 0 ? pieces : new List<string> { address };
    }

    private static void ExpandByFormulas(object excel, string originAddress, List<string> hits)
    {
        var snapshot = hits.ToArray();
        foreach (var address in snapshot)
        {
            var widened = ExtendRow(excel, originAddress, address);
            if (!string.IsNullOrEmpty(widened)) hits.Add(widened);
        }

        var sheets = new Dictionary<string, ColumnSpan>(StringComparer.OrdinalIgnoreCase);
        foreach (var address in hits)
        {
            var normalized = DependentsTree.NormalizeAddress(address, originAddress);
            var parsed = TraceUtils.ParseWorksheetScopedAddress(normalized);
            if (parsed == null || !TraceUtils.IsRangeAddress(parsed.RangeAddress)) continue;
            var parts = parsed.RangeAddress.Split(':');
            var start = TraceUtils.ParseCellAddress(parts[0]);
            var end = TraceUtils.ParseCellAddress(parts[parts.Length - 1]);
            if (start == null || end == null) continue;
            if (!sheets.TryGetValue(parsed.WorksheetName, out var span))
            {
                span = new ColumnSpan
                {
                    Min = Math.Min(start.Value.Col, end.Value.Col),
                    Max = Math.Max(start.Value.Col, end.Value.Col)
                };
                sheets[parsed.WorksheetName] = span;
            }
            else
            {
                span.Min = Math.Min(span.Min, Math.Min(start.Value.Col, end.Value.Col));
                span.Max = Math.Max(span.Max, Math.Max(start.Value.Col, end.Value.Col));
            }
        }

        foreach (var sheet in sheets)
        {
            DiscoverRows(excel, originAddress, sheet.Key, sheet.Value.Min, sheet.Value.Max, hits);
        }
    }

    private static string ExtendRow(object excel, string originAddress, string address)
    {
        var normalized = DependentsTree.NormalizeAddress(address, originAddress);
        var parsed = TraceUtils.ParseWorksheetScopedAddress(normalized);
        if (parsed == null || !TraceUtils.IsRangeAddress(parsed.RangeAddress)) return "";
        var parts = parsed.RangeAddress.Split(':');
        var start = TraceUtils.ParseCellAddress(parts[0]);
        var end = TraceUtils.ParseCellAddress(parts[parts.Length - 1]);
        if (start == null || end == null || start.Value.Row != end.Value.Row) return "";

        var row = start.Value.Row;
        var minCol = Math.Min(start.Value.Col, end.Value.Col);
        var maxCol = Math.Max(start.Value.Col, end.Value.Col);
        var guard = 0;
        while (minCol > 0 && guard < 64 &&
               CellReferencesOrigin(excel, parsed.WorksheetName, row, minCol - 1, originAddress))
        {
            minCol--;
            guard++;
        }

        guard = 0;
        while (maxCol < 16383 && guard < 64 &&
               CellReferencesOrigin(excel, parsed.WorksheetName, row, maxCol + 1, originAddress))
        {
            maxCol++;
            guard++;
        }

        var widened = "'" + parsed.WorksheetName.Replace("'", "''") + "'!" +
                      TraceUtils.ColumnLetters(minCol) + (row + 1).ToString(CultureInfo.InvariantCulture);
        if (minCol != maxCol)
        {
            widened += ":" + TraceUtils.ColumnLetters(maxCol) + (row + 1).ToString(CultureInfo.InvariantCulture);
        }

        return widened;
    }

    private static void DiscoverRows(
        object excel,
        string originAddress,
        string sheet,
        int minCol,
        int maxCol,
        List<string> hits)
    {
        var probe = RangeOf(excel, "'" + sheet.Replace("'", "''") + "'!A1");
        if (probe == null) return;
        object? worksheet;
        try { worksheet = Get(probe, "Worksheet"); }
        catch { return; }
        if (worksheet == null) return;

        object? used;
        try { used = Get(worksheet, "UsedRange"); }
        catch { return; }
        if (used == null) return;

        int firstRow;
        int rowCount;
        try
        {
            firstRow = Convert.ToInt32(Get(used, "Row"), CultureInfo.InvariantCulture);
            var rows = Get(used, "Rows");
            rowCount = Convert.ToInt32(Get(rows!, "Count"), CultureInfo.InvariantCulture);
        }
        catch
        {
            return;
        }

        if (rowCount <= 0) return;
        if (rowCount > 2000) rowCount = 2000;
        var lastRow = firstRow + rowCount - 1;
        var width = maxCol - minCol + 1;
        if (width <= 0 || width > 40) return;

        var blockAddress = "'" + sheet.Replace("'", "''") + "'!" +
                           TraceUtils.ColumnLetters(minCol) + firstRow.ToString(CultureInfo.InvariantCulture) + ":" +
                           TraceUtils.ColumnLetters(maxCol) + lastRow.ToString(CultureInfo.InvariantCulture);
        var block = RangeOf(excel, blockAddress);
        if (block == null) return;
        object? formulas;
        try { formulas = Get(block, "Formula"); }
        catch { return; }

        var originNorm = DependentsTree.NormalizeAddress(originAddress, originAddress);
        for (var row = firstRow; row <= lastRow; row++)
        {
            for (var col = minCol; col <= maxCol; col++)
            {
                var formula = FormulaAt(formulas, row - firstRow, col - minCol, width, rowCount);
                if (!DependentsTree.FormulaReferencesOrigin(formula, sheet, originAddress)) continue;
                var cell = "'" + sheet.Replace("'", "''") + "'!" +
                           TraceUtils.ColumnLetters(col) + row.ToString(CultureInfo.InvariantCulture);
                if (TraceUtils.AddressesReferToSameRange(cell, originNorm)) continue;
                hits.Add(cell);
            }
        }
    }

    private static string FormulaAt(object? formulas, int rowOffset, int columnOffset, int width, int height)
    {
        if (formulas is string text) return text;
        if (formulas is Array array && array.Rank == 1)
        {
            var index = array.GetLowerBound(0) + columnOffset;
            if (index > array.GetUpperBound(0)) return "";
            return array.GetValue(index) as string ?? "";
        }

        if (formulas is Array grid && grid.Rank == 2)
        {
            var row = grid.GetLowerBound(0) + rowOffset;
            var column = grid.GetLowerBound(1) + columnOffset;
            if (row > grid.GetUpperBound(0) || column > grid.GetUpperBound(1)) return "";
            return grid.GetValue(row, column) as string ?? "";
        }

        if (width == 1 && height == 1 && formulas != null) return formulas.ToString() ?? "";
        return "";
    }

    private static bool CellReferencesOrigin(object excel, string sheet, int rowIndex, int columnIndex, string originAddress)
    {
        var address = "'" + sheet.Replace("'", "''") + "'!" +
                      TraceUtils.ColumnLetters(columnIndex) + (rowIndex + 1).ToString(CultureInfo.InvariantCulture);
        var range = RangeOf(excel, address);
        if (range == null) return false;
        string? formula = null;
        try { formula = Get(range, "Formula") as string; }
        catch { formula = null; }
        if (string.IsNullOrEmpty(formula))
        {
            try { formula = Get(range, "FormulaLocal") as string; }
            catch { formula = null; }
        }

        return DependentsTree.FormulaReferencesOrigin(formula, sheet, originAddress);
    }

    private sealed class ColumnSpan
    {
        public int Min { get; set; }
        public int Max { get; set; }
    }

    private static string AddressOf(object range)
    {
        string? text = null;
        try
        {
            text = InvokeAddress(range, external: true);
        }
        catch
        {
            try { text = InvokeAddress(range, external: false); }
            catch { /* FormatRange is the fallback */ }
        }

        if (string.IsNullOrWhiteSpace(text)) return "";
        if (TraceUtils.ParseWorksheetScopedAddress(text!) != null) return text!.Trim();

        var sheet = "";
        try
        {
            var worksheet = Get(range, "Worksheet");
            if (worksheet != null) sheet = Get(worksheet, "Name") as string ?? "";
        }
        catch
        {
            sheet = "";
        }

        if (string.IsNullOrEmpty(sheet)) return text!.Trim();
        return "'" + sheet.Replace("'", "''") + "'!" + text!.Trim();
    }

    private static string? InvokeAddress(object range, bool external)
    {
        var property = BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance;
        try
        {
            var value = range.GetType().InvokeMember(
                "Address",
                property,
                null,
                range,
                new object[] { false, false, 1, external, Missing.Value });
            return value as string;
        }
        catch
        {
            var value = range.GetType().InvokeMember(
                "Address",
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                range,
                new object[] { false, false, 1, external });
            return value as string;
        }
    }

    private static string FormatRange(object? range)
    {
        if (range == null) return "";
        var sheetObj = Get(range, "Worksheet");
        var sheet = sheetObj == null ? "" : Get(sheetObj, "Name") as string ?? "";
        var row = Convert.ToInt32(Get(range, "Row"), CultureInfo.InvariantCulture);
        var column = Convert.ToInt32(Get(range, "Column"), CultureInfo.InvariantCulture);
        var rows = Get(range, "Rows");
        var columns = Get(range, "Columns");
        var rowCount = Convert.ToInt32(Get(rows!, "Count"), CultureInfo.InvariantCulture);
        var columnCount = Convert.ToInt32(Get(columns!, "Count"), CultureInfo.InvariantCulture);
        if (row <= 0 || column <= 0 || rowCount <= 0 || columnCount <= 0) return "";

        var start = TraceUtils.ColumnLetters(column - 1) + row.ToString(CultureInfo.InvariantCulture);
        var end = TraceUtils.ColumnLetters(column + columnCount - 2) +
                  (row + rowCount - 1).ToString(CultureInfo.InvariantCulture);
        var a1 = start == end ? start : start + ":" + end;
        if (string.IsNullOrEmpty(sheet)) return a1;
        return "'" + sheet.Replace("'", "''") + "'!" + a1;
    }

    private static object? RangeOf(object excel, string address)
    {
        try
        {
            return excel.GetType().InvokeMember(
                "Range",
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                excel,
                new object[] { address, Missing.Value });
        }
        catch
        {
            return null;
        }
    }

    private static object? Get(object target, string name)
    {
        return target.GetType().InvokeMember(
            name,
            BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
            null,
            target,
            null);
    }

    private static object? Invoke(object target, string name, params object[] args)
    {
        return target.GetType().InvokeMember(
            name,
            BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
            null,
            target,
            args.Length == 0 ? null : args);
    }

    private static void Set(object target, string name, object? value)
    {
        target.GetType().InvokeMember(
            name,
            BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
            null,
            target,
            new[] { value });
    }

    private static object? Item(object target, int index)
    {
        try
        {
            return target.GetType().InvokeMember(
                "Item",
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                target,
                new object[] { index });
        }
        catch
        {
            return Invoke(target, "Item", index);
        }
    }
}
