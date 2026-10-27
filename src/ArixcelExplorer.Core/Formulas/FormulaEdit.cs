namespace ArixcelExplorer.Core.Formulas;

public static class FormulaEdit
{
    public static string Normalize(string? formula)
    {
        var text = (formula ?? "").Trim();
        if (text.Length == 0) return "";
        if (text.StartsWith("=", System.StringComparison.Ordinal) ||
            text.StartsWith("{=", System.StringComparison.Ordinal))
        {
            return text;
        }

        return "=" + text;
    }
}
