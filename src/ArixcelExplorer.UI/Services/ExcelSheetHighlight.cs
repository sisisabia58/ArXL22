using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ArixcelExplorer.Core.Settings;

namespace ArixcelExplorer.UI.Services;

/// <summary>
/// Selects a whole dependent or precedent range and paints it. Excel's Goto
/// during host navigation leaves only the top-left cell selected. Closing the
/// window restores the interior colors this highlighter changed.
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
            var range = RangeOf(excel, address!);
            if (range == null) return;
            ForEachCell(range, RememberBaseline);
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
        if (string.IsNullOrWhiteSpace(address)) return;
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
            PreserveSaved(excel, () =>
            {
                var range = RangeOf(excel, address!);
                if (range == null) return;
                try
                {
                    var sheet = Get(range, "Worksheet");
                    if (sheet != null) Invoke(sheet, "Activate");
                }
                catch
                {
                    // Select still moves to the range when the sheet is already active.
                }

                Invoke(range, "Select");
                ScrollWithoutCollapsing(excel, range);
                var color = ExcelOleColor.FromHex(string.IsNullOrWhiteSpace(hex) ? "#C6E5B3" : hex);
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
            });
        }
        catch
        {
            // Highlight is optional; the tree selection still stands.
        }
    }

    public void CloseHighlight()
    {
        _closed = true;
        RestoreActive();
        RestoreAll(_baseline);
        RestoreAll(_owned);
        try
        {
            System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                new Action(() => RestoreAll(_baseline)),
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
                    Apply(fill);
                }
            });
        }
        catch
        {
            foreach (var fill in fills)
            {
                Apply(fill);
            }
        }
    }

    private void RestoreActive()
    {
        foreach (var fill in _active)
        {
            Apply(fill);
        }

        _active.Clear();
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
        var key = KeyOf(cell);
        if (string.IsNullOrEmpty(key)) return null;
        try
        {
            var interior = Get(cell, "Interior");
            if (interior == null) return null;
            var fill = ReadFill(cell, interior, key);
            _owned.Add(fill);
            return fill;
        }
        catch
        {
            return null;
        }
    }

    private void RememberBaseline(object cell)
    {
        var key = KeyOf(cell);
        if (string.IsNullOrEmpty(key) || !_seen.Add(key)) return;
        try
        {
            var interior = Get(cell, "Interior");
            if (interior == null) return;
            _baseline.Add(ReadFill(cell, interior, key));
        }
        catch
        {
            _seen.Remove(key);
        }
    }

    private static CellFill ReadFill(object cell, object interior, string key)
    {
        object? colorIndex = null;
        object? pattern = null;
        object? color = null;
        try { colorIndex = Get(interior, "ColorIndex"); } catch { /* optional */ }
        try { pattern = Get(interior, "Pattern"); } catch { /* optional */ }
        try { color = Get(interior, "Color"); } catch { /* optional */ }
        return new CellFill
        {
            Cell = cell,
            Key = key,
            NoFill = IsNone(colorIndex),
            ColorIndex = colorIndex,
            Pattern = pattern,
            Color = color
        };
    }

    private static void Apply(CellFill fill)
    {
        try
        {
            var interior = Get(fill.Cell, "Interior");
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
            return Convert.ToInt32(Convert.ToDouble(value, CultureInfo.InvariantCulture)) == XlNone;
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

    private sealed class CellFill
    {
        public object Cell { get; set; } = null!;
        public string Key { get; set; } = "";
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
