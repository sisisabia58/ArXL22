using System.Collections.Generic;
using ArixcelExplorer.Core.Audit;

namespace ArixcelExplorer.Core.Audit;

public enum FlowCellRole
{
    Unknown,
    Input,
    Calculation,
    Output
}

public sealed class FlowCellInfo
{
    public string Address { get; set; } = "";
    public bool HasFormula { get; set; }
    public bool HasPrecedentsInSelection { get; set; }
    public bool HasDependentsInSelection { get; set; }
    public AutoColorCategory ColorCategory { get; set; }
}

public static class CalculationFlow
{
    public static FlowCellRole ClassifyCell(FlowCellInfo cell)
    {
        if (!cell.HasFormula)
        {
            return FlowCellRole.Input;
        }

        if (!cell.HasPrecedentsInSelection && cell.HasDependentsInSelection)
        {
            return FlowCellRole.Input;
        }

        if (cell.HasPrecedentsInSelection && !cell.HasDependentsInSelection)
        {
            return FlowCellRole.Output;
        }

        return FlowCellRole.Calculation;
    }

    public static FlowCellRole[][] ClassifyGrid(IReadOnlyList<IReadOnlyList<FlowCellInfo>> grid)
    {
        var result = new FlowCellRole[grid.Count][];
        for (var r = 0; r < grid.Count; r++)
        {
            result[r] = new FlowCellRole[grid[r].Count];
            for (var c = 0; c < grid[r].Count; c++)
            {
                result[r][c] = ClassifyCell(grid[r][c]);
            }
        }

        return result;
    }

    public static string GetFlowColor(FlowCellRole role) => role switch
    {
        FlowCellRole.Input => "#FFF2CC",
        FlowCellRole.Calculation => "#E2EFDA",
        FlowCellRole.Output => "#DDEBF7",
        _ => "#FFFFFF"
    };
}
