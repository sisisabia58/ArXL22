using System;
using System.Text.RegularExpressions;

namespace ArixcelExplorer.Core.Audit;

public enum AutoColorCategory
{
    None,
    Input,
    Formula,
    WorksheetLink,
    WorkbookLink,
    External,
    Hyperlink,
    PartialInput
}

public sealed class AutoColorCell
{
    public string? Formula { get; set; }
    public object? Value { get; set; }
    public string? NumberFormat { get; set; }
    public bool HasHyperlink { get; set; }
}

public sealed class AutoColorPalette
{
    public string Input { get; set; } = "#0000FF";
    public string Formula { get; set; } = "#000000";
    public string WorksheetLink { get; set; } = "#008000";
    public string WorkbookLink { get; set; } = "#CC99FF";
    public string External { get; set; } = "#00B0F0";
    public string Hyperlink { get; set; } = "#FF8000";
    public string PartialInput { get; set; } = "#800000";
    public string Inconsistent { get; set; } = "#FF0000";
    public string Consistent { get; set; } = "#00B050";
}

public static class AutoColor
{
    private static readonly string[] CommonFunctions =
    {
        "SUM", "AVERAGE", "COUNT", "LEFT", "RIGHT", "MID", "ROUND"
    };

    public static AutoColorCategory ClassifyCell(AutoColorCell cell)
    {
        var formula = cell.Formula;
        if (IsFormulaCell(formula))
        {
            if (IsPartialInputFormula(cell)) return AutoColorCategory.PartialInput;
            if (IsWorkbookLinkFormula(formula!)) return AutoColorCategory.WorkbookLink;
            if (IsWorksheetLinkFormula(formula!)) return AutoColorCategory.WorksheetLink;
            if (IsExternalReferenceFormula(formula!)) return AutoColorCategory.External;
            if (IsInputCell(cell)) return AutoColorCategory.Input;
            return AutoColorCategory.Formula;
        }

        if (cell.HasHyperlink) return AutoColorCategory.Hyperlink;
        if (IsInputCell(cell)) return AutoColorCategory.Input;
        return AutoColorCategory.None;
    }

    public static AutoColorCategory[][] ClassifyGrid(AutoColorCell[][] cells)
    {
        var result = new AutoColorCategory[cells.Length][];
        for (var r = 0; r < cells.Length; r++)
        {
            result[r] = new AutoColorCategory[cells[r].Length];
            for (var c = 0; c < cells[r].Length; c++)
            {
                result[r][c] = ClassifyCell(cells[r][c]);
            }
        }

        return result;
    }

    public static string? GetColorForCategory(AutoColorCategory category, AutoColorPalette palette)
    {
        return category switch
        {
            AutoColorCategory.Input => palette.Input,
            AutoColorCategory.Formula => palette.Formula,
            AutoColorCategory.WorksheetLink => palette.WorksheetLink,
            AutoColorCategory.WorkbookLink => palette.WorkbookLink,
            AutoColorCategory.External => palette.External,
            AutoColorCategory.Hyperlink => palette.Hyperlink,
            AutoColorCategory.PartialInput => palette.PartialInput,
            _ => null
        };
    }

    public static bool IsFormulaCell(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula)) return false;
        var trimmed = formula.Trim();
        return trimmed.StartsWith("=", StringComparison.Ordinal) ||
               (trimmed.StartsWith("{=", StringComparison.Ordinal) && trimmed.EndsWith("}", StringComparison.Ordinal));
    }

    public static bool IsInputCell(AutoColorCell cell)
    {
        var value = cell.Value;
        var formula = cell.Formula;
        if (value == null || value is string { Length: 0 }) return false;
        if (value is string) return false;
        if (IsFormulaCell(formula))
        {
            if (!HasCellReference(formula!)) return true;
            return IsOnlyNumbersAndOperators(formula!);
        }

        return true;
    }

    public static bool IsOnlyNumbersAndOperators(string formula)
    {
        var normalized = NormalizeFormula(formula);
        return Regex.IsMatch(normalized, @"^[-+*/\d\s.,()]*$");
    }

    public static bool IsWorkbookLinkFormula(string formula) =>
        formula.IndexOf('!') >= 0 && !Regex.IsMatch(formula, @"\[[^\]]+\]");

    public static bool IsWorksheetLinkFormula(string formula) =>
        HasCellReference(formula) && formula.IndexOf('!') < 0 && !Regex.IsMatch(formula, @"\[[^\]]+\]");

    public static bool IsExternalReferenceFormula(string formula)
    {
        var upper = formula.ToUpperInvariant();
        return Regex.IsMatch(formula, @"\[[^\]]+\]") ||
               upper.Contains("WEBSERVICE") ||
               upper.Contains("ODBC") ||
               upper.Contains("SQL");
    }

    public static bool IsPartialInputFormula(AutoColorCell cell)
    {
        var formula = cell.Formula;
        if (!IsFormulaCell(formula)) return false;
        if (cell.Value is string) return false;
        if (IsOnlyNumbersAndOperators(formula!)) return false;
        if (!HasCellReference(formula!)) return false;

        var candidate = NormalizeFormula(formula!);
        candidate = Regex.Replace(candidate, @"(\[[^\]]+\])?'?[^!]+!'?", "SHEET_REF!");
        foreach (var func in CommonFunctions)
        {
            candidate = Regex.Replace(candidate, func, "", RegexOptions.IgnoreCase);
        }

        candidate = candidate.Replace("$", "").Replace("%", "");
        candidate = Regex.Replace(candidate, @"[$]?[A-Za-z]+[$]?[0-9]+", "");
        return Regex.IsMatch(candidate, @"[0-9]+");
    }

    private static bool HasCellReference(string formula) =>
        Regex.IsMatch(NormalizeFormula(formula), @"[$]?[A-Za-z]+[$]?[0-9]+|R[0-9]*C[0-9]*");

    private static string NormalizeFormula(string formula)
    {
        var trimmed = formula.Trim();
        if (trimmed.StartsWith("{=", StringComparison.Ordinal) && trimmed.EndsWith("}", StringComparison.Ordinal))
        {
            return trimmed.Substring(2, trimmed.Length - 3);
        }

        if (trimmed.StartsWith("=", StringComparison.Ordinal))
        {
            return trimmed.Substring(1);
        }

        return trimmed;
    }
}
