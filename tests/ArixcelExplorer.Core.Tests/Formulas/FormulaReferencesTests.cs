using System.Linq;
using ArixcelExplorer.Core.Formulas;
using Xunit;

namespace ArixcelExplorer.Core.Tests.Formulas;

public sealed class FormulaReferencesTests
{
    [Fact]
    public void TokenizeFormulaReferences_finds_local_and_sheet_qualified_refs()
    {
        var segments = FormulaReferences.TokenizeFormulaReferences("=SUM(A1,'Sheet 2'!B2)", "Sheet1");
        var refs = segments.OfType<FormulaReferenceSegment>().ToList();
        Assert.Equal(2, refs.Count);
        Assert.Contains("A1", refs[0].Text);
        Assert.Contains("Sheet 2", refs[1].Text);
    }

    [Fact]
    public void Parse_creates_tree_for_sum_function()
    {
        var tree = FormulaAstParser.Parse("=SUM(A1:A5,B1)", "'S'!C1", "15");
        Assert.NotEmpty(tree.Children);
        Assert.Equal("SUM", tree.Children[0].Label);
        Assert.True(tree.Children[0].Children.Count >= 1);
    }
}
