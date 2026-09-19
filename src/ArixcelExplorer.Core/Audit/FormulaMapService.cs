using System.Collections.Generic;

namespace ArixcelExplorer.Core.Audit;

public sealed class FormulaMapCellResult
{
    public string Address { get; set; } = "";
    public string? FillColor { get; set; }
    public string Category { get; set; } = "";
}

public static class FormulaMapService
{
    public static IReadOnlyList<FormulaMapCellResult> BuildMap(
        AutoColorCell[][] cells,
        string[][] addresses,
        FormulaConsistencyMark[][]? consistencyMarks,
        AutoColorPalette palette)
    {
        var categories = AutoColor.ClassifyGrid(cells);
        var results = new List<FormulaMapCellResult>();

        for (var r = 0; r < cells.Length; r++)
        {
            for (var c = 0; c < cells[r].Length; c++)
            {
                string? color = null;
                var category = categories[r][c].ToString();

                if (consistencyMarks != null &&
                    cells[r][c].Formula != null &&
                    consistencyMarks[r][c] == FormulaConsistencyMark.Inconsistent)
                {
                    color = palette.Inconsistent;
                    category = "Inconsistent";
                }
                else if (consistencyMarks != null &&
                         consistencyMarks[r][c] == FormulaConsistencyMark.Consistent)
                {
                    color = palette.Consistent;
                    category = "Consistent";
                }
                else
                {
                    color = AutoColor.GetColorForCategory(categories[r][c], palette);
                }

                if (color == null) continue;

                results.Add(new FormulaMapCellResult
                {
                    Address = addresses[r][c],
                    FillColor = color,
                    Category = category
                });
            }
        }

        return results;
    }
}
