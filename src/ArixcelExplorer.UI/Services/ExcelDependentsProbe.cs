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
            try { Invoke(origin, "Select"); } catch { /* DirectDependents does not need the selection */ }
            try { ReadAreas(origin, areas); } catch { /* no same-sheet dependents */ }
            try
            {
                if (ShowDependents(origin))
                {
                    WalkArrows(excel, originAddress!, arrows);
                }
            }
            catch
            {
                // Cross-sheet arrows are optional; same-sheet areas still stand.
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
        foreach (var piece in Expand(deps))
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

    private static void WalkArrows(object excel, string originAddress, List<string> arrows)
    {
        var originNorm = DependentsTree.NormalizeAddress(originAddress, originAddress);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var duplicateArrows = 0;
        for (var arrow = 1; arrow <= 1024; arrow++)
        {
            var added = 0;
            var linkWorked = false;
            for (var link = 1; link <= 16; link++)
            {
                object? dest;
                try
                {
                    var cell = RangeOf(excel, originAddress);
                    if (cell == null) return;
                    // NavigateArrow follows arrows on the active cell. Reselect the
                    // source before each arrow; ScreenUpdating is off for this walk.
                    try { Invoke(cell, "Select"); } catch { /* the range call below still addresses it */ }
                    dest = Invoke(cell, "NavigateArrow", false, arrow, link);
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
                foreach (var piece in Expand(dest))
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

    private static List<string> Expand(object range)
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
                        var formatted = FormatRange(area);
                        if (!string.IsNullOrEmpty(formatted)) pieces.Add(formatted);
                    }

                    if (pieces.Count > 0) return pieces;
                }
            }
        }
        catch
        {
            // Fall through to the range's own rectangle.
        }

        try
        {
            var single = FormatRange(range);
            if (!string.IsNullOrEmpty(single)) return new List<string> { single };
        }
        catch
        {
            // Unreadable COM range.
        }

        return new List<string>();
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
