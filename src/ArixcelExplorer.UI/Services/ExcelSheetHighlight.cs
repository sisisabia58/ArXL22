using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.UI.Services;

/// <summary>
/// Selects a whole dependent or precedent range and paints it. Excel's Goto
/// during host navigation leaves only the top-left cell selected. Closing the
/// window restores the interior colors this highlighter changed.
/// </summary>
internal sealed class ExcelSheetHighlight
{
    private readonly Dictionary<string, int> _original = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _painted = new List<string>();

    public void Show(string? address, string? hex)
    {
        Clear();
        if (string.IsNullOrWhiteSpace(address)) return;
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
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
        }
        catch
        {
            // Highlight is optional; the tree selection still stands.
        }
    }

    public void Clear()
    {
        if (_painted.Count == 0) return;
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
            foreach (var address in _painted)
            {
                if (!_original.TryGetValue(address, out var color)) continue;
                var cell = RangeOf(excel, address);
                if (cell == null) continue;
                try
                {
                    var interior = Get(cell, "Interior");
                    if (interior != null) Set(interior, "Color", color);
                }
                catch
                {
                    // A deleted sheet should not block the rest of the restore.
                }
            }
        }
        catch
        {
            // Excel is gone; nothing left to restore.
        }

        _painted.Clear();
        _original.Clear();
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
                var address = ExternalAddress(cell);
                if (string.IsNullOrEmpty(address)) continue;
                try
                {
                    var interior = Get(cell, "Interior");
                    if (interior == null) continue;
                    if (!_original.ContainsKey(address))
                    {
                        _original[address] = Convert.ToInt32(Convert.ToDouble(Get(interior, "Color"), CultureInfo.InvariantCulture));
                    }

                    Set(interior, "Color", color);
                    _painted.Add(address);
                }
                catch
                {
                    // Skip a cell whose interior cannot be read.
                }
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

    private static string ExternalAddress(object cell)
    {
        try
        {
            var value = cell.GetType().InvokeMember(
                "Address",
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                cell,
                new object[] { false, false, 1, true, Missing.Value });
            var text = value as string;
            if (!string.IsNullOrWhiteSpace(text)) return text!.Trim();
        }
        catch
        {
            // Fall through to a sheet-qualified address built from the indexes.
        }

        try
        {
            var sheetObj = Get(cell, "Worksheet");
            var sheet = sheetObj == null ? "" : Get(sheetObj, "Name") as string ?? "";
            var row = Count(Get(cell, "Row"));
            var column = Count(Get(cell, "Column"));
            var a1 = TraceUtils.ColumnLetters(column - 1) + row.ToString(CultureInfo.InvariantCulture);
            if (string.IsNullOrEmpty(sheet)) return a1;
            return "'" + sheet.Replace("'", "''") + "'!" + a1;
        }
        catch
        {
            return "";
        }
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
