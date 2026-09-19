using ArixcelExplorer.Core.Audit;
using Xunit;

namespace ArixcelExplorer.Core.Tests.Audit;

public sealed class AutoColorTests
{
    [Fact]
    public void ClassifyCell_marks_numeric_constant_as_input()
    {
        var category = AutoColor.ClassifyCell(new AutoColorCell { Value = 42 });
        Assert.Equal(AutoColorCategory.Input, category);
    }

    [Fact]
    public void AnalyzeHorizontalFormulaConsistency_marks_copy_errors()
    {
        var rows = new[]
        {
            new[]
            {
                new FormulaConsistencyCell { IsFormula = true, FormulaR1C1 = "RC[1]" },
                new FormulaConsistencyCell { IsFormula = true, FormulaR1C1 = "RC[2]" }
            }
        };

        var marks = FormulaConsistency.AnalyzeHorizontalFormulaConsistency(rows);
        Assert.Equal(FormulaConsistencyMark.Inconsistent, marks[0][0]);
    }
}
