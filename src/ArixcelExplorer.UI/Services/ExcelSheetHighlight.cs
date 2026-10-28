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

    private readonly List<PaintedCell> _painted = new List<PaintedCell>();
    private bool _closed;

    public void Show(string? address, string? hex)
    {
        if (_closed) return;
        Clear();
        if (_closed || string.IsNullOrWhiteSpace(address)) return;
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
                Paint(range, ExcelOleColor.FromHex(string.IsNullOrWhiteSpace(hex) ? "#C6E5B3" : hex));
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
        Clear();
    }

    public void Clear()
    {
        if (_painted.Count == 0) return;
        var pending = _painted.ToArray();
        _painted.Clear();
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
            PreserveSaved(excel, () =>
            {
                foreach (var painted in pending)
                {
                    Restore(painted);
                }
            });
        }
        catch
        {
            foreach (var painted in pending)
            {
                Restore(painted);
            }
        }
    }

    private static void Restore(PaintedCell painted)
    {
        try
        {
            var interior = Get(painted.Cell, "Interior");
            if (interior == null) return;
            if (painted.NoFill)
            {
                Set(interior, "Pattern", XlNone);
                return;
            }

            if (painted.Pattern != null) Set(interior, "Pattern", painted.Pattern);
            if (painted.Color != null) Set(interior, "Color", painted.Color);
        }
        catch
        {
            // Keep going so one stale cell does not leave the rest green.
        }
    }

    private void Paint(object range, int color)
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

        if (rows <= 0 || columns <= 0) return;
        if (rows > 400) rows = 400;
        if (columns > 80) columns = 80;
        for (var row = 1; row <= rows; row++)
        {
            for (var column = 1; column <= columns; column++)
            {
                object? cell;
                try { cell = Item(range, row, column); }
                catch { continue; }
                if (cell == null) continue;
                try
                {
                    var interior = Get(cell, "Interior");
                    if (interior == null) continue;
                    var snapshot = Snapshot(interior);
                    Set(interior, "Color", color);
                    _painted.Add(new PaintedCell { Cell = cell, NoFill = snapshot.NoFill, Pattern = snapshot.Pattern, Color = snapshot.Color });
                }
                catch
                {
                    // Skip a cell whose interior cannot be read. Do not paint it.
                }
            }
        }
    }

    private static PaintedCell Snapshot(object interior)
    {
        object? colorIndex = null;
        object? pattern = null;
        object? color = null;
        try { colorIndex = Get(interior, "ColorIndex"); } catch { /* optional */ }
        try { pattern = Get(interior, "Pattern"); } catch { /* optional */ }
        try { color = Get(interior, "Color"); } catch { /* optional */ }
        var noFill = IsNone(colorIndex) || IsNone(pattern);
        return new PaintedCell { NoFill = noFill, Pattern = pattern, Color = color };
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

    private sealed class PaintedCell
    {
        public object Cell { get; set; } = null!;
        public object? Color { get; set; }
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
