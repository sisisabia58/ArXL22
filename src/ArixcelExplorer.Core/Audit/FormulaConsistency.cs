using System.Collections.Generic;

namespace ArixcelExplorer.Core.Audit;

public enum FormulaConsistencyMark
{
    None,
    Consistent,
    Inconsistent
}

public sealed class FormulaConsistencyCell
{
    public bool IsFormula { get; set; }
    public string? FormulaR1C1 { get; set; }
}

public static class FormulaConsistency
{
    public static HashSet<string> CollectAdjacentEqualFormulas(IReadOnlyList<IReadOnlyList<FormulaConsistencyCell>> rows)
    {
        var consistentFormulas = new HashSet<string>();
        foreach (var row in rows)
        {
            for (var c = 0; c < row.Count - 1; c++)
            {
                var current = row[c];
                var right = row[c + 1];
                if (!current.IsFormula || !right.IsFormula) continue;

                var currentFormula = current.FormulaR1C1 ?? "";
                var rightFormula = right.FormulaR1C1 ?? "";
                if (currentFormula.Length > 0 && currentFormula == rightFormula)
                {
                    consistentFormulas.Add(currentFormula);
                }
            }
        }

        return consistentFormulas;
    }

    public static FormulaConsistencyMark[][] AnalyzeHorizontalFormulaConsistency(
        IReadOnlyList<IReadOnlyList<FormulaConsistencyCell>> rows)
    {
        var consistentFormulas = CollectAdjacentEqualFormulas(rows);
        var result = new FormulaConsistencyMark[rows.Count][];

        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            result[r] = new FormulaConsistencyMark[row.Count];
            for (var c = 0; c < row.Count; c++)
            {
                var cell = row[c];
                if (!cell.IsFormula)
                {
                    result[r][c] = FormulaConsistencyMark.None;
                    continue;
                }

                var currentFormula = cell.FormulaR1C1 ?? "";
                var right = c + 1 < row.Count ? row[c + 1] : null;
                if (right?.IsFormula == true)
                {
                    var rightFormula = right.FormulaR1C1 ?? "";
                    result[r][c] = currentFormula.Length > 0 && currentFormula == rightFormula
                        ? FormulaConsistencyMark.Consistent
                        : FormulaConsistencyMark.Inconsistent;
                }
                else
                {
                    result[r][c] = currentFormula.Length > 0 && consistentFormulas.Contains(currentFormula)
                        ? FormulaConsistencyMark.Consistent
                        : FormulaConsistencyMark.None;
                }
            }
        }

        return result;
    }
}
