using System;
using System.Text.RegularExpressions;

namespace ArixcelExplorer.Core.Tracing;

public enum TraceDirection
{
    Precedents,
    Dependents
}

public sealed class ParsedTraceAddress
{
    public string WorksheetName { get; set; } = "";
    public string RangeAddress { get; set; } = "";
}

public static class TraceUtils
{
    public const int DefaultTraceMaxDepth = 10;
    public const int MaxTraceMaxDepth = 20;
    public const int DefaultTraceSafetyLimit = 500;
    public const int MaxTraceSafetyLimit = 5000;
    public const int MaxTraceRows = DefaultTraceSafetyLimit;

    public static int SanitizeTraceDepth(object? raw)
    {
        if (raw is not int value || !IsFinite(value))
        {
            return DefaultTraceMaxDepth;
        }

        var normalized = value;
        if (normalized < 1) return 1;
        if (normalized > MaxTraceMaxDepth) return MaxTraceMaxDepth;
        return normalized;
    }

    public static int SanitizeTraceSafetyLimit(object? raw)
    {
        if (raw is not int value || !IsFinite(value))
        {
            return DefaultTraceSafetyLimit;
        }

        var normalized = value;
        if (normalized < 1) return 1;
        if (normalized > MaxTraceSafetyLimit) return MaxTraceSafetyLimit;
        return normalized;
    }

    public static string FormatTraceValue(object? value)
    {
        if (value == null) return "";
        if (value is string s) return s;
        if (value is double or float or int or long or decimal or bool)
        {
            return Convert.ToString(value) ?? "";
        }

        return value.ToString() ?? "";
    }

    public static string FormatTraceFormula(object? value)
    {
        if (value is not string formula) return "";
        return formula.StartsWith("=", StringComparison.Ordinal) ? formula : "";
    }

    public static string BuildTraceCellKey(string worksheetName, int rowIndex, int columnIndex)
    {
        return $"{worksheetName}!R{rowIndex}C{columnIndex}";
    }

    public static ParsedTraceAddress? ParseWorksheetScopedAddress(string address)
    {
        var trimmed = address.Trim();
        var bang = trimmed.LastIndexOf('!');
        if (bang <= 0 || bang >= trimmed.Length - 1) return null;

        var worksheetName = trimmed.Substring(0, bang).Trim();
        var rangeAddress = trimmed.Substring(bang + 1).Trim();
        if (string.IsNullOrEmpty(worksheetName) || string.IsNullOrEmpty(rangeAddress)) return null;

        if (worksheetName.StartsWith("'") && worksheetName.EndsWith("'") && worksheetName.Length >= 2)
        {
            worksheetName = worksheetName.Substring(1, worksheetName.Length - 2).Replace("''", "'");
        }

        if (worksheetName.StartsWith("[", StringComparison.Ordinal))
        {
            var close = worksheetName.IndexOf(']');
            if (close > 0 && close < worksheetName.Length - 1)
            {
                worksheetName = worksheetName.Substring(close + 1);
            }
        }

        if (string.IsNullOrEmpty(worksheetName)) return null;
        return new ParsedTraceAddress { WorksheetName = worksheetName, RangeAddress = rangeAddress };
    }

    public static (int Row, int Col)? ParseCellAddress(string address)
    {
        var match = Regex.Match(address.Trim(), @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$");
        if (!match.Success) return null;
        return (int.Parse(match.Groups[2].Value) - 1, ColumnIndex(match.Groups[1].Value));
    }

    public static string ColumnLetters(int columnIndex)
    {
        var index = columnIndex + 1;
        var letters = "";
        while (index > 0)
        {
            var rem = (index - 1) % 26;
            letters = (char)('A' + rem) + letters;
            index = (index - 1) / 26;
        }

        return letters;
    }

    public static int ColumnIndex(string letters)
    {
        var result = 0;
        foreach (var c in letters.ToUpperInvariant())
        {
            result = result * 26 + (c - 'A' + 1);
        }

        return result - 1;
    }

    private static bool IsFinite(int value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
