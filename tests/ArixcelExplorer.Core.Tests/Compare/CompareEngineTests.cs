using ArixcelExplorer.Core.Compare;
using Xunit;

namespace ArixcelExplorer.Core.Tests.Compare;

public sealed class CompareEngineTests
{
    [Fact]
    public void DiffGrids_reports_value_mismatch()
    {
        var left = new[]
        {
            new[] { new CompareCellSnapshot { Address = "A1", Value = 1 } }
        };
        var right = new[]
        {
            new[] { new CompareCellSnapshot { Address = "A1", Value = 2 } }
        };

        var diffs = CompareEngine.DiffGrids(left, right, "Left", "Right");
        Assert.Single(diffs);
        Assert.Equal("ValueOrFormula", diffs[0].DiffKind);
    }
}
