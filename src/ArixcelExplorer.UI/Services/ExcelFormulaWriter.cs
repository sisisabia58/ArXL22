using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ArixcelExplorer.UI.Services;

/// <summary>
/// Writes a formula back to the running Excel instance. The installed host
/// does not have to be rebuilt for the formula pane to commit.
/// </summary>
internal static class ExcelFormulaWriter
{
    public static bool TryWrite(string? address, string? formula)
    {
        if (string.IsNullOrWhiteSpace(address) || string.IsNullOrWhiteSpace(formula)) return false;
        var target = address!;
        var text = formula!;
        try
        {
            var excel = Marshal.GetActiveObject("Excel.Application");
            var range = excel.GetType().InvokeMember(
                "Range",
                BindingFlags.GetProperty,
                null,
                excel,
                new object[] { target, Missing.Value });
            if (range == null) return false;
            if (TrySet(range, "Formula", text)) return true;
            return TrySet(range, "FormulaLocal", text);
        }
        catch
        {
            return false;
        }
    }

    private static bool TrySet(object range, string property, string formula)
    {
        try
        {
            range.GetType().InvokeMember(
                property,
                BindingFlags.SetProperty,
                null,
                range,
                new object[] { formula });
            return true;
        }
        catch
        {
            return false;
        }
    }
}
