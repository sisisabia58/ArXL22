using System.Text.RegularExpressions;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.Core.Formulas;

public sealed class ValidationListSource
{
    public string Label { get; set; } = "";
    public string Location { get; set; } = "";
    public bool IsNamedRange { get; set; }
}

public static class ValidationListParser
{
    private static readonly Regex NamedRangePattern = new(
        @"^[A-Za-z_\\][A-Za-z0-9_.]*$",
        RegexOptions.CultureInvariant);

    private static readonly Regex R1C1Pattern = new(
        @"^R(\[-?\d+\]|\d+)C(\[-?\d+\]|\d+)(?::R(\[-?\d+\]|\d+)C(\[-?\d+\]|\d+))?$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static ValidationListSource? TryParse(string? formula1, string contextWorksheetName)
    {
        _ = contextWorksheetName;
        var text = NormalizeFormula1(formula1);
        if (text.Length == 0) return null;

        if (text.IndexOf(',') >= 0 && text.IndexOf(':') < 0 && text.IndexOf('!') < 0)
        {
            return null;
        }

        var scoped = TraceUtils.ParseWorksheetScopedAddress(text);
        if (scoped != null)
        {
            return new ValidationListSource
            {
                Label = text,
                Location = text,
                IsNamedRange = false
            };
        }

        if (LooksLikeA1Range(text) || IsR1C1(text))
        {
            return new ValidationListSource
            {
                Label = text,
                Location = text,
                IsNamedRange = false
            };
        }

        if (NamedRangePattern.IsMatch(text))
        {
            return new ValidationListSource
            {
                Label = text,
                Location = text,
                IsNamedRange = true
            };
        }

        return null;
    }

    public static bool IsR1C1(string? location)
    {
        if (string.IsNullOrWhiteSpace(location)) return false;
        var text = location!.Trim();
        var range = TraceUtils.ParseWorksheetScopedAddress(text)?.RangeAddress ?? text;
        return R1C1Pattern.IsMatch(range.Trim());
    }

    private static string NormalizeFormula1(string? formula1)
    {
        if (string.IsNullOrWhiteSpace(formula1)) return "";
        var text = formula1!.Trim().Trim('"').Trim();
        if (text.StartsWith("=")) text = text.Substring(1).Trim();
        return text.Trim('"').Trim();
    }

    private static bool LooksLikeA1Range(string text)
    {
        return Regex.IsMatch(
            text.Trim(),
            @"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }
}
