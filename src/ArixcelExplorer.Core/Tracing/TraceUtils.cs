using System;
using System.Globalization;
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
    public const string ArrayValue = "<array>";

    private static readonly Regex CvErrToken = new(@"-21468262\d{2}", RegexOptions.Compiled);

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
        if (value is string s) return FormatTraceText(s);
        if (TryFormatExcelError(value, out var errorName)) return errorName;
        if (value is double d) return FormatGeneralNumber(d);
        if (value is float f) return FormatGeneralNumber(f);
        if (value is decimal m) return FormatGeneralNumber(m);
        if (value is int or long or short or byte)
        {
            return FormatGroupedInteger(Convert.ToDouble(value, CultureInfo.InvariantCulture));
        }

        if (value is Array array)
        {
            return FormatArrayValue(array);
        }

        var raw = value.ToString() ?? "";
        if (raw.Contains("System.Object")) return "";
        return FormatTraceText(raw);
    }

    /// <summary>
    /// A failed Excel Evaluate comes back as a CVErr HRESULT (for example -2146826273 = #VALUE!).
    /// That integer must not be shown next to a real cell value.
    /// </summary>
    public static bool TryFormatExcelError(object? value, out string name)
    {
        name = "";
        long? code = value switch
        {
            int i => i,
            long l => l,
            short sh => sh,
            byte by => by,
            double d when IsIntegral(d) && d >= int.MinValue && d <= int.MaxValue => (long)Math.Round(d),
            float f when IsIntegral(f) && f >= int.MinValue && f <= int.MaxValue => (long)Math.Round(f),
            decimal m when m == decimal.Truncate(m) && m >= int.MinValue && m <= int.MaxValue => (long)m,
            string s when long.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) => parsed,
            _ => null
        };
        if (code == null) return false;
        name = ExcelErrorName(code.Value);
        return name.Length > 0;
    }

    public static bool IsUsableFormulaResult(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text!.Trim();
        if (IsExcelErrorName(trimmed)) return false;
        if (TryFormatExcelError(trimmed, out _)) return false;
        return true;
    }

    public static bool IsMultiCellRange(string? address)
    {
        if (string.IsNullOrWhiteSpace(address)) return false;
        var text = address!;
        if (!IsRangeAddress(text)) return false;
        var parsed = ParseWorksheetScopedAddress(text);
        var range = (parsed?.RangeAddress ?? text.Trim()).Replace("$", "");
        var parts = range.Split(':');
        return parts.Length == 2 &&
               !string.Equals(parts[0], parts[1], StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Small integers stay plain (6, 4). Larger numbers pick up the current culture's thousands separator.
    /// </summary>
    public static string FormatGeneralNumber(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value)) return "";
        if (TryFormatExcelError(value, out var errorName)) return errorName;
        var rounded = Math.Round(value, 12, MidpointRounding.AwayFromZero);
        var nearest = Math.Round(rounded, MidpointRounding.AwayFromZero);
        if (Math.Abs(rounded - nearest) < 1e-9)
        {
            return FormatGroupedInteger(nearest);
        }

        return rounded.ToString("#,0.############", CultureInfo.CurrentCulture);
    }

    public static string FormatGeneralNumber(decimal value)
    {
        if (TryFormatExcelError(value, out var errorName)) return errorName;
        if (value == decimal.Truncate(value))
        {
            return FormatGroupedInteger((double)decimal.Truncate(value));
        }

        return value.ToString("#,0.#############################", CultureInfo.CurrentCulture);
    }

    private static string FormatTraceText(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        if (text == ArrayValue) return text;
        var trimmed = text.Trim();
        if (IsExcelErrorName(trimmed)) return trimmed;
        if (TryFormatExcelError(trimmed, out var onlyError)) return onlyError;

        if (!CvErrToken.IsMatch(trimmed)) return FormatNumericText(trimmed);

        var stripped = CvErrToken.Replace(trimmed, " ");
        stripped = Regex.Replace(stripped, @"[\s,;]+", " ").Trim();
        if (stripped.Length == 0) return "";
        return FormatNumericText(stripped);
    }

    private static string FormatNumericText(string text)
    {
        if (TryParseFlexibleNumber(text, out var number)) return FormatGeneralNumber(number);
        var normalized = NormalizeNumericText(text);
        if (!string.Equals(normalized, text, StringComparison.Ordinal) &&
            TryParseFlexibleNumber(normalized, out number))
        {
            return FormatGeneralNumber(number);
        }

        return normalized;
    }

    private static string FormatArrayValue(Array array)
    {
        var count = 0;
        string? single = null;
        if (!CollectArrayValues(array, ref count, ref single)) return ArrayValue;
        if (count <= 0) return "";
        if (count == 1) return single ?? "";
        return ArrayValue;
    }

    private static bool CollectArrayValues(Array array, ref int count, ref string? single)
    {
        foreach (var item in array)
        {
            if (item is Array nested)
            {
                if (!CollectArrayValues(nested, ref count, ref single)) return false;
                continue;
            }

            if (item == null || item is DBNull || TryFormatExcelError(item, out _)) continue;
            var formatted = FormatTraceValue(item);
            if (formatted.Length == 0 || formatted == ArrayValue || IsExcelErrorName(formatted)) continue;
            count += 1;
            if (count == 1) single = formatted;
            if (count >= 2) return false;
        }

        return true;
    }

    private static string FormatGroupedInteger(double value)
    {
        var nearest = Math.Round(value, MidpointRounding.AwayFromZero);
        if (Math.Abs(nearest) < 1000)
        {
            return nearest.ToString("0", CultureInfo.InvariantCulture);
        }

        return nearest.ToString("#,0", CultureInfo.CurrentCulture);
    }

    private static bool TryParseFlexibleNumber(string text, out double value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text.Trim();
        if (Regex.IsMatch(trimmed, @"^-?\d+$"))
        {
            return double.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        var match = Regex.Match(trimmed, @"^(-?\d+)([.,])(\d+)$");
        if (!match.Success) return false;
        var fraction = match.Groups[3].Value.TrimEnd('0');
        var canonical = fraction.Length == 0
            ? match.Groups[1].Value
            : match.Groups[1].Value + "." + fraction;
        return double.TryParse(canonical, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    }

    private static bool IsIntegral(double value) =>
        !double.IsNaN(value) && !double.IsInfinity(value) &&
        Math.Abs(value - Math.Round(value, MidpointRounding.AwayFromZero)) < 1e-6;

    private static bool IsExcelErrorName(string text) =>
        text is "#NULL!" or "#DIV/0!" or "#VALUE!" or "#REF!" or "#NAME?" or "#NUM!" or "#N/A" or "#GETTING_DATA";

    private static string ExcelErrorName(long code) => code switch
    {
        -2146826288 => "#NULL!",
        -2146826281 => "#DIV/0!",
        -2146826273 => "#VALUE!",
        -2146826265 => "#REF!",
        -2146826259 => "#NAME?",
        -2146826252 => "#NUM!",
        -2146826246 => "#N/A",
        -2146826245 => "#GETTING_DATA",
        _ => ""
    };

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
                   string.Equals(StripAbsolute(parsedLeft.RangeAddress), StripAbsolute(parsedRight.RangeAddress), StringComparison.OrdinalIgnoreCase);
        }

        var leftRange = StripAbsolute(parsedLeft?.RangeAddress ?? leftText);
        var rightRange = StripAbsolute(parsedRight?.RangeAddress ?? rightText);
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
