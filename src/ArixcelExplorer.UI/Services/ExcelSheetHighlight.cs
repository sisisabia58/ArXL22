using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.UI.Services;

/// <summary>
/// Selects the precedent areas and paints those cells. A multi-area address
/// such as K21 and V65 is selected and painted as separate areas, not as the
/// rectangle between them. Closing the window restores the interior colors
/// this highlighter changed.
/// </summary>
internal sealed class ExcelSheetHighlight
{
    private const int XlNone = -4142;

    private readonly List<CellFill> _baseline = new List<CellFill>();
    private readonly List<CellFill> _owned = new List<CellFill>();
    private readonly HashSet<string> _seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private readonly List<CellFill> _active = new List<CellFill>();
    private bool _closed;

    public void Capture(string? address)
    {
        if (_closed || string.IsNullOrWhiteSpace(address)) return;
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
            foreach (var part in TraceUtils.SplitAreas(address))
            {
                var range = RangeOf(excel, part);
                if (range == null) continue;
                ForEachCell(range, RememberBaseline);
            }
        }
        catch
        {
            // Missing cells are skipped. Painted cells still restore from what was captured.
        }
    }

    public void Show(string? address, string? hex)
    {
        if (_closed) return;
        RestoreActive();
        var parts = TraceUtils.SplitAreas(address);
        if (parts.Count == 0) return;
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
            PreserveSaved(excel, () =>
            {
                var ranges = new List<object>();
                foreach (var part in parts)
                {
                    var range = RangeOf(excel, part);
                    if (range != null) ranges.Add(range);
                }

                if (ranges.Count == 0) return;
                var selection = ranges[0];
                for (var i = 1; i < ranges.Count; i++)
                {
                    try
                    {
                        var union = Invoke(excel, "Union", selection, ranges[i]);
                        if (union != null) selection = union;
                    }
                    catch
                    {
                        // Paint the areas we could resolve even if Union is unavailable.
                    }
                }

                try
                {
                    var sheet = Get(ranges[0], "Worksheet");
                    if (sheet != null) Invoke(sheet, "Activate");
                }
                catch
                {
                    // Select still moves to the range when the sheet is already active.
                }

                Invoke(selection, "Select");
                ScrollWithoutCollapsing(excel, ranges[0]);
                var color = ExcelOleColor.FromHex(string.IsNullOrWhiteSpace(hex) ? "#C6E5B3" : hex);
                foreach (var range in ranges)
                {
                    ForEachCell(range, cell =>
                    {
                        try
                        {
                            var fill = Find(_baseline, cell) ?? Own(cell);
                            if (fill == null) return;
                            var interior = Get(cell, "Interior");
                            if (interior == null) return;
                            Set(interior, "Color", color);
                            _active.Add(fill);
                        }
                        catch
                        {
                            // Leave a cell unpainted when its interior cannot be changed.
                        }
                    });
                }
            });
        }
        catch
        {
            // Highlight is optional; the tree selection still stands.
        }
    }

    public void SuspendPaints() => _closed = true;

    public void CloseHighlight()
    {
        _closed = true;
        RestoreActive();
        RestoreAll(_baseline);
        RestoreAll(_owned);
        ClearTransientSheetState();
        try
        {
            System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                new Action(() =>
                {
                    RestoreAll(_baseline);
                    RestoreAll(_owned);
                    ClearTransientSheetState();
                }),
                System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }
        catch
        {
            // The synchronous restore already ran.
        }
    }

    public void Clear()
    {
        RestoreActive();
    }

    private void RestoreAll(List<CellFill> fills)
    {
        if (fills.Count == 0) return;
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
            PreserveSaved(excel, () =>
            {
                foreach (var fill in fills)
                {
                    Apply(excel, fill);
                }
            });
        }
        catch
        {
            foreach (var fill in fills)
            {
                Apply(null, fill);
            }
        }
    }

    private void RestoreActive()
    {
        object? excel = null;
        try { excel = Marshal.GetActiveObject("Excel.Application"); }
        catch { excel = null; }

        foreach (var fill in _active)
        {
            Apply(excel, fill);
        }

        _active.Clear();
    }

    private static void ClearTransientSheetState()
    {
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
            PreserveSaved(excel, () =>
            {
                try { Set(excel, "CutCopyMode", false); }
                catch { /* no copy marquee to clear */ }

                try
                {
                    var sheet = Get(excel, "ActiveSheet");
                    if (sheet != null) Invoke(sheet, "ClearArrows");
                }
                catch
                {
                    // Arrows are optional.
                }

                try
                {
                    var selection = Get(excel, "Selection");
                    if (selection == null) return;
                    var count = Count(Get(selection, "Count"));
                    if (count <= 1) return;
                    var active = Get(excel, "ActiveCell");
                    if (active != null) Invoke(active, "Select");
                }
                catch
                {
                    // Leave a non-range selection alone.
                }
            });
        }
        catch
        {
            // Excel is not reachable from this thread.
        }
    }

    private CellFill? Find(List<CellFill> fills, object cell)
    {
        var key = KeyOf(cell);
        if (string.IsNullOrEmpty(key)) return null;
        foreach (var fill in fills)
        {
            if (string.Equals(fill.Key, key, StringComparison.OrdinalIgnoreCase)) return fill;
        }

        return null;
    }

    private CellFill? Own(object cell)
    {
        var existing = Find(_owned, cell);
        if (existing != null) return existing;
        var fill = ReadFill(cell);
        if (fill == null) return null;
        _owned.Add(fill);
        return fill;
    }

    private void RememberBaseline(object cell)
    {
        var fill = ReadFill(cell);
        if (fill == null || string.IsNullOrEmpty(fill.Key) || !_seen.Add(fill.Key)) return;
        _baseline.Add(fill);
    }

    private static CellFill? ReadFill(object cell)
    {
        var key = KeyOf(cell);
        if (string.IsNullOrEmpty(key)) return null;
        try
        {
            var interior = Get(cell, "Interior");
            if (interior == null) return null;
            object? colorIndex = null;
            object? pattern = null;
            object? color = null;
            try { colorIndex = Get(interior, "ColorIndex"); } catch { /* optional */ }
            try { pattern = Get(interior, "Pattern"); } catch { /* optional */ }
            try { color = Get(interior, "Color"); } catch { /* optional */ }
            var sheetObj = Get(cell, "Worksheet");
            var sheet = sheetObj == null ? "" : Get(sheetObj, "Name") as string ?? "";
            return new CellFill
            {
                Cell = cell,
                Key = key,
                Sheet = sheet,
                Row = Count(Get(cell, "Row")),
                Column = Count(Get(cell, "Column")),
                NoFill = IsNone(colorIndex),
                ColorIndex = colorIndex,
                Pattern = pattern,
                Color = color
            };
        }
        catch
        {
            return null;
        }
    }

    private static void Apply(object? excel, CellFill fill)
    {
        try
        {
            var cell = ResolveCell(excel, fill) ?? fill.Cell;
            var interior = Get(cell, "Interior");
            if (interior == null) return;
            if (fill.NoFill)
            {
                try { Set(interior, "ColorIndex", XlNone); } catch { /* try pattern */ }
                try { Set(interior, "Pattern", XlNone); } catch { /* already cleared */ }
                return;
            }

            if (fill.Pattern != null)
            {
                try { Set(interior, "Pattern", fill.Pattern); } catch { /* color still restores */ }
            }

            if (fill.ColorIndex != null)
            {
                try { Set(interior, "ColorIndex", fill.ColorIndex); } catch { /* use Color */ }
            }

            if (fill.Color != null)
            {
                try { Set(interior, "Color", fill.Color); } catch { /* best effort */ }
            }
        }
        catch
        {
            // One stale cell must not skip the rest of the restore.
        }
    }

    private static bool IsNone(object? value)
    {
        if (value == null) return false;
        try
        {
            return InteriorFill.IsNoFill(Convert.ToInt32(Convert.ToDouble(value, CultureInfo.InvariantCulture)));
        }
        catch
        {
            return false;
        }
    }

    private static void PreserveSaved(object excel, Action action)
    {
        object? workbook = null;
        var wasSaved = false;
        var knowSaved = false;
        try
        {
            workbook = Get(excel, "ActiveWorkbook");
            if (workbook != null)
            {
                wasSaved = Convert.ToBoolean(Get(workbook, "Saved"), CultureInfo.InvariantCulture);
                knowSaved = true;
            }
        }
        catch
        {
            knowSaved = false;
        }

        try
        {
            action();
        }
        finally
        {
            if (knowSaved && wasSaved && workbook != null)
            {
                try { Set(workbook, "Saved", true); }
                catch { /* the restore still happened */ }
            }
        }
    }

    private static void ScrollWithoutCollapsing(object excel, object range)
    {
        try
        {
            var window = Get(excel, "ActiveWindow");
            if (window == null) return;
            var row = Count(Get(range, "Row"));
            var column = Count(Get(range, "Column"));
            var visible = Get(window, "VisibleRange");
            if (visible == null)
            {
                Set(window, "ScrollRow", row);
                Set(window, "ScrollColumn", column);
                return;
            }

            var top = Count(Get(visible, "Row"));
            var left = Count(Get(visible, "Column"));
            var height = Count(Get(visible, "Rows"));
            var width = Count(Get(visible, "Columns"));
            if (row < top || row >= top + height) Set(window, "ScrollRow", row);
            if (column < left || column >= left + width) Set(window, "ScrollColumn", column);
        }
        catch
        {
            // The range is already selected even if it stays off screen.
        }
    }

    private static void ForEachCell(object range, Action<object> action)
    {
        object? areas = null;
        var areaCount = 0;
        try
        {
            areas = Get(range, "Areas");
            areaCount = areas == null ? 0 : Count(Get(areas, "Count"));
        }
        catch
        {
            areaCount = 0;
        }

        if (areaCount > 1 && areas != null)
        {
            for (var i = 1; i <= areaCount; i++)
            {
                object? area = null;
                try { area = AreaItem(areas, i); }
                catch { continue; }
                if (area != null) ForEachRectangle(area, action);
            }

            return;
        }

        ForEachRectangle(range, action);
    }

    private static void ForEachRectangle(object range, Action<object> action)
    {
        int rows;
        int columns;
        try
        {
            rows = Count(Get(range, "Rows"));
            columns = Count(Get(range, "Columns"));
        }
        catch
        {
            return;
        }

        if (rows <= 0 || columns <= 0 || rows > 400 || columns > 80) return;
        for (var row = 1; row <= rows; row++)
        {
            for (var column = 1; column <= columns; column++)
            {
                object? cell;
                try { cell = Item(range, row, column); }
                catch { continue; }
                if (cell != null) action(cell);
            }
        }
    }

    private static string KeyOf(object cell)
    {
        try
        {
            var sheetObj = Get(cell, "Worksheet");
            var sheet = sheetObj == null ? "" : Get(sheetObj, "Name") as string ?? "";
            var row = Count(Get(cell, "Row"));
            var column = Count(Get(cell, "Column"));
            if (row <= 0 || column <= 0) return "";
            return sheet + "!" + row.ToString(CultureInfo.InvariantCulture) + ":" + column.ToString(CultureInfo.InvariantCulture);
        }
        catch
        {
            return "";
        }
    }

    private static object? ResolveCell(object? excel, CellFill fill)
    {
        if (excel == null || string.IsNullOrEmpty(fill.Sheet) || fill.Row <= 0 || fill.Column <= 0) return null;
        try
        {
            var sheets = Get(excel, "Worksheets");
            if (sheets == null) return null;
            var sheet = Invoke(sheets, "Item", fill.Sheet);
            if (sheet == null) return null;
            var cells = Get(sheet, "Cells");
            if (cells == null) return null;
            return Item(cells, fill.Row, fill.Column);
        }
        catch
        {
            return null;
        }
    }

    private sealed class CellFill
    {
        public object Cell { get; set; } = null!;
        public string Key { get; set; } = "";
        public string Sheet { get; set; } = "";
        public int Row { get; set; }
        public int Column { get; set; }
        public object? Color { get; set; }
        public object? ColorIndex { get; set; }
        public object? Pattern { get; set; }
        public bool NoFill { get; set; }
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

    private static int Count(object? value)
    {
        if (value == null) return 0;
        try
        {
            var count = Get(value, "Count");
            if (count != null) return Convert.ToInt32(count, CultureInfo.InvariantCulture);
        }
        catch
        {
            // Value is already a number, such as Range.Row.
        }

        return Convert.ToInt32(Convert.ToDouble(value, CultureInfo.InvariantCulture));
    }

    private static object? AreaItem(object areas, int index)
    {
        try
        {
            return areas.GetType().InvokeMember(
                "Item",
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                areas,
                new object[] { index });
        }
        catch
        {
            return areas.GetType().InvokeMember(
                "Item",
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                areas,
                new object[] { index });
        }
    }

    private static object? Item(object range, int row, int column)
    {
        try
        {
            return range.GetType().InvokeMember(
                "Item",
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                range,
                new object[] { row, column });
        }
        catch
        {
            return range.GetType().InvokeMember(
                "Item",
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                range,
                new object[] { row, column });
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

    private static void Set(object target, string name, object value)
    {
        target.GetType().InvokeMember(
            name,
            BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
            null,
            target,
            new object[] { value });
    }
}

internal static class ExcelSheetHighlightColors
{
    private static string? _dependent;
    private static string? _precedent;

    public static string Dependent => _dependent ?? (_dependent = Read(true));

    public static string Precedent => _precedent ?? (_precedent = Read(false));

    private static string Read(bool dependent)
    {
        try
        {
            var options = OptionsStore.Load();
            var hex = dependent ? options.DependentHighlight : options.PrecedentHighlight;
            if (ExcelOleColor.IsValidHex(hex)) return hex;
        }
        catch
        {
            // Use the same defaults the options window ships with.
        }

        return dependent ? "#C6E5B3" : "#7FDBFF";
    }
}
