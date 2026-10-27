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
        if (value == null || value is DBNull) return "";
        if (value is bool b) return b ? "TRUE" : "FALSE";
        if (value is string s) return NormalizeNumericText(s);
        if (value is double d) return FormatGeneralNumber(d);
        if (value is float f) return FormatGeneralNumber(f);
        if (value is decimal m) return FormatGeneralNumber(m);
        if (value is int or long or short or byte)
        {
            return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "";
        }

        if (value is Array array)
        {
            var flat = Flatten(array);
            if (flat.Count == 0) return "";
            if (flat.Count < 4) return string.Join(", ", flat);
            return flat.Count + " values";
        }

        var raw = value.ToString() ?? "";
        if (raw.Contains("System.Object")) return "";
        return NormalizeNumericText(raw);
    }

    /// <summary>
    /// Arixcel shows general-format numbers (6, 2, 4), not locale fixed decimals (6,00).
    /// </summary>
    public static string FormatGeneralNumber(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value)) return "";
        var rounded = Math.Round(value, 12, MidpointRounding.AwayFromZero);
        if (Math.Abs(rounded - Math.Round(rounded, MidpointRounding.AwayFromZero)) < 1e-9)
        {
            return Math.Round(rounded, MidpointRounding.AwayFromZero)
                .ToString("0", System.Globalization.CultureInfo.InvariantCulture);
        }

        return rounded.ToString("0.############", System.Globalization.CultureInfo.CurrentCulture);
    }

    public static string FormatGeneralNumber(decimal value)
    {
        if (value == decimal.Truncate(value))
        {
            return decimal.Truncate(value).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return value.ToString("0.#############################", System.Globalization.CultureInfo.CurrentCulture);
    }

    public static string NormalizeNumericText(string text)
    {
        if (string.IsNullOrEmpty(text)) return text ?? "";
        var trimmed = text.Trim();
        var match = Regex.Match(trimmed, @"^(-?\d+)([.,])(\d+)$");
        if (!match.Success) return text;
        var fraction = match.Groups[3].Value.TrimEnd('0');
        if (fraction.Length == 0) return match.Groups[1].Value;
        return match.Groups[1].Value + match.Groups[2].Value + fraction;
    }

    /// <summary>
    /// Smallest single range covering every address. D4 and E4 on one sheet become 'Sheet'!D4:E4.
    /// </summary>
    public static string BoundingAddress(System.Collections.Generic.IEnumerable<string?> addresses)
    {
        string? sheet = null;
        var minRow = int.MaxValue;
        var minCol = int.MaxValue;
        var maxRow = -1;
        var maxCol = -1;
        var any = false;
        foreach (var address in addresses)
        {
            if (string.IsNullOrWhiteSpace(address)) continue;
            var parsed = ParseWorksheetScopedAddress(address!);
            var range = parsed?.RangeAddress ?? address!.Trim();
            if (!IsRangeAddress(range)) continue;
            var worksheet = parsed?.WorksheetName;
            if (sheet == null)
            {
                sheet = worksheet;
            }
            else if (!string.Equals(sheet, worksheet, StringComparison.OrdinalIgnoreCase))
            {
                return "";
            }

            foreach (var part in range.Split(':'))
            {
                var cell = ParseCellAddress(part);
                if (cell == null) continue;
                any = true;
                minRow = Math.Min(minRow, cell.Value.Row);
                maxRow = Math.Max(maxRow, cell.Value.Row);
                minCol = Math.Min(minCol, cell.Value.Col);
                maxCol = Math.Max(maxCol, cell.Value.Col);
            }
        }

        if (!any) return "";
        var start = ColumnLetters(minCol) + (minRow + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var end = ColumnLetters(maxCol) + (maxRow + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var a1 = start == end ? start : start + ":" + end;
        if (string.IsNullOrEmpty(sheet)) return a1;
        return $"'{sheet}'!{a1}";
    }

    private static System.Collections.Generic.List<string> Flatten(Array array)
    {
        var items = new System.Collections.Generic.List<string>();
        foreach (var item in array)
        {
            if (item is Array nested)
            {
                items.AddRange(Flatten(nested));
            }
            else
            {
                var formatted = FormatTraceValue(item);
                if (formatted.Length > 0) items.Add(formatted);
            }
        }

        return items;
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

    public static bool AddressesReferToSameRange(string? left, string? right)
    {
        if (left is null || right is null) return false;
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;

        var leftText = left.Trim();
        var rightText = right.Trim();
        if (string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase)) return true;

        var parsedLeft = ParseWorksheetScopedAddress(leftText);
        var parsedRight = ParseWorksheetScopedAddress(rightText);
        if (parsedLeft != null && parsedRight != null)
        {
            return string.Equals(parsedLeft.WorksheetName, parsedRight.WorksheetName, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(parsedLeft.RangeAddress, parsedRight.RangeAddress, StringComparison.OrdinalIgnoreCase);
        }

        var leftRange = parsedLeft?.RangeAddress ?? leftText;
        var rightRange = parsedRight?.RangeAddress ?? rightText;
        return string.Equals(leftRange, rightRange, StringComparison.OrdinalIgnoreCase);
    }

    public static string QualifyAddress(string address, string contextAddress)
    {
        if (string.IsNullOrWhiteSpace(address)) return "";
        var trimmed = address.Trim();
        if (ParseWorksheetScopedAddress(trimmed) != null) return trimmed;
        if (!CanQualifyAsAddress(trimmed)) return "";
        var context = ParseWorksheetScopedAddress(contextAddress);
        if (context == null) return trimmed;
        return $"'{context.WorksheetName}'!{trimmed}";
    }

    public static string FormatExplorerLocation(string? location, string originAddress)
    {
        if (string.IsNullOrWhiteSpace(location)) return "";
        var qualified = QualifyAddress(location ?? "", originAddress);
        var parsed = ParseWorksheetScopedAddress(string.IsNullOrEmpty(qualified) ? location.Trim() : qualified);
        if (parsed == null || !IsRangeAddress(parsed.RangeAddress)) return "";

        var shortRange = StripAbsolute(parsed.RangeAddress);
        var origin = ParseWorksheetScopedAddress(originAddress);
        if (origin != null &&
            string.Equals(origin.WorksheetName, parsed.WorksheetName, StringComparison.OrdinalIgnoreCase))
        {
            return shortRange;
        }

        return $"'{parsed.WorksheetName}'!{shortRange}";
    }

    public static string VisibleAnchor(string? address)
    {
        if (string.IsNullOrWhiteSpace(address)) return "";
        var trimmed = address.Trim();
        var parsed = ParseWorksheetScopedAddress(trimmed);
        var range = parsed?.RangeAddress ?? trimmed;
        if (!IsRangeAddress(range)) return "";
        var first = range.Split(':')[0];
        return StripAbsolute(first);
    }

    public static bool IsRangeAddress(string address)
    {
        if (string.IsNullOrWhiteSpace(address)) return false;
        var trimmed = address.Trim();
        var parsed = ParseWorksheetScopedAddress(trimmed);
        var range = parsed?.RangeAddress ?? trimmed;
        var parts = range.Split(':');
        if (parts.Length is not (1 or 2)) return false;
        foreach (var part in parts)
        {
            if (ParseCellAddress(part) == null) return false;
        }

        return true;
    }

    private static bool CanQualifyAsAddress(string trimmed)
    {
        if (IsRangeAddress(trimmed)) return true;
        return Regex.IsMatch(trimmed, @"^[A-Za-z_\\][A-Za-z0-9_.]*$");
    }

    private static string StripAbsolute(string range) => range.Replace("$", "");

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
