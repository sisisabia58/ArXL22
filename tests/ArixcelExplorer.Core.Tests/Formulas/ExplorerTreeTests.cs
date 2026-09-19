using System.Linq;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;
using Xunit;

namespace ArixcelExplorer.Core.Tests.Formulas;

public sealed class ExplorerTreeTests
{
    [Fact]
    public void FlattenVisible_includes_origin_root_as_first_row()
    {
        var tree = FormulaAstParser.Parse("=SUM(A1:A5,B1)", "'S'!C1", "15");
        var rows = FormulaAstParser.FlattenVisible(tree);

        Assert.Equal(FormulaNodeKind.Root, rows[0].Kind);
        Assert.Equal("'S'!C1", rows[0].Label);
        Assert.Equal("'S'!C1", rows[0].Location);
        Assert.Contains(rows, node => node.Label == "SUM");
    }

    [Fact]
    public void CycleExpandAll_collapses_descendants_then_restores_them()
    {
        var tree = FormulaAstParser.Parse("=SUM(A1:A5,B1)", "'S'!C1", "15");
        Assert.True(tree.IsExpanded);
        Assert.True(tree.Children[0].IsExpanded);

        var expanded = FormulaAstParser.CycleExpandAll(tree, currentlyExpanded: true);
        Assert.False(expanded);
        Assert.True(tree.IsExpanded);
        Assert.False(tree.Children[0].IsExpanded);

        var afterCollapse = FormulaAstParser.FlattenVisible(tree);
        Assert.Equal("'S'!C1", afterCollapse[0].Label);
        Assert.Contains(afterCollapse, node => node.Label == "SUM");
        Assert.DoesNotContain(afterCollapse, node => node.Label == "A1:A5");

        expanded = FormulaAstParser.CycleExpandAll(tree, currentlyExpanded: false);
        Assert.True(expanded);
        Assert.True(tree.Children[0].IsExpanded);

        var afterExpand = FormulaAstParser.FlattenVisible(tree);
        Assert.Contains(afterExpand, node => node.Label == "A1:A5");
        Assert.Contains(afterExpand, node => node.Label == "B1");
    }

    [Fact]
    public void QualifyAddress_prefixes_sheet_from_origin()
    {
        var qualified = TraceUtils.QualifyAddress("A1:A5", "'Sheet 1'!C1");
        Assert.Equal("'Sheet 1'!A1:A5", qualified);
        Assert.Equal("'Other'!B2", TraceUtils.QualifyAddress("'Other'!B2", "'Sheet 1'!C1"));
        Assert.Equal("", TraceUtils.QualifyAddress(" ", "'Sheet 1'!C1"));
    }

    [Fact]
    public void ExcelOleColor_converts_hex_to_bgr()
    {
        Assert.True(ExcelOleColor.IsValidHex("#F4C2C2"));
        Assert.False(ExcelOleColor.IsValidHex("red"));
        Assert.Equal(0xC2C2F4, ExcelOleColor.FromHex("#F4C2C2"));
        Assert.Equal(0xC16305, ExcelOleColor.FromHex("#0563C1"));
        Assert.Equal(0xFFFFFF, ExcelOleColor.FromHex("nope"));
    }
}
