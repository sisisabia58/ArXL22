using System.Collections.Generic;

namespace ArixcelExplorer.Core.Formulas;

public sealed class FormulaHighlightSpan
{
    public int Start { get; set; }
    public int Length { get; set; }
    public bool IsOrigin { get; set; }
}

public static class FormulaHighlight
{
    public const string CyanHex = "#7FDBFF";

    public static IReadOnlyList<FormulaHighlightSpan> ForSelectedRows(IEnumerable<FormulaHighlightSpan> selected)
    {
        var result = new List<FormulaHighlightSpan>();
        foreach (var span in selected)
        {
            if (span.IsOrigin || span.Length <= 0 || span.Start < 0) continue;
            result.Add(span);
        }

        return result;
    }
}
