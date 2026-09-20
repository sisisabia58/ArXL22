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
    public void CollapseAll_leaves_function_row_with_plus_glyph()
    {
        var tree = FormulaAstParser.Parse("=SUMIFS(A1:A5,B1:B5,C1)", "'S'!G15", "1");
        FormulaAstParser.CollapseAll(tree);
        var rows = FormulaAstParser.FlattenVisible(tree);
        var function = Assert.Single(rows, node => node.Label == "SUMIFS");

        Assert.True(function.Children.Count > 0);
        Assert.False(function.IsExpanded);
        Assert.Equal("+", ExplorerTreeChrome.Glyph(function.Children.Count > 0, function.IsExpanded));
        Assert.DoesNotContain(rows, node => node.Info == "sum_range");
    }

    [Theory]
    [InlineData(true, true, "-")]
    [InlineData(true, false, "+")]
    [InlineData(false, false, "")]
    public void Glyph_matches_folder_tree_plus_minus(bool hasChildren, bool expanded, string expected)
    {
        Assert.Equal(expected, ExplorerTreeChrome.Glyph(hasChildren, expanded));
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

    [Fact]
    public void FormatTraceValue_does_not_emit_system_object_array()
    {
        object[,] block =
        {
            { 10.5, 20d },
            { 30d, 40d }
        };

        Assert.Equal("4 values", TraceUtils.FormatTraceValue(block));
        Assert.Equal("15", TraceUtils.FormatTraceValue(15d));
        Assert.Equal("", TraceUtils.FormatTraceValue(null));
        Assert.Equal("hello", TraceUtils.FormatTraceValue("hello"));
    }

    [Fact]
    public void FormatTraceValue_formats_one_by_n_array_as_joined_preview()
    {
        object[] row = { 1d, 2d, 3d };
        Assert.Equal("1, 2, 3", TraceUtils.FormatTraceValue(row));
    }

    [Fact]
    public void Default_precedent_highlight_is_cyan()
    {
        Assert.Equal("#7FDBFF", ArixcelOptions.Default.PrecedentHighlight);
    }

    [Fact]
    public void FunctionArgInfo_labels_if_and_sumifs_arguments()
    {
        Assert.Equal("logical_test", FunctionArgInfo.LabelFor("IF", 0));
        Assert.Equal("value_if_true", FunctionArgInfo.LabelFor("IF", 1));
        Assert.Equal("value_if_false", FunctionArgInfo.LabelFor("IF", 2));
        Assert.Equal("sum_range", FunctionArgInfo.LabelFor("SUMIFS", 0));
        Assert.Equal("criteria_range1", FunctionArgInfo.LabelFor("SUMIFS", 1));
        Assert.Equal("criteria1", FunctionArgInfo.LabelFor("SUMIFS", 2));
        Assert.Equal("", FunctionArgInfo.LabelFor("SUM", 0));
    }

    [Fact]
    public void Parse_records_display_spans_for_if_logical_test()
    {
        const string formula = "=IF(I$3>4,A1,B1)";
        var tree = FormulaAstParser.Parse(formula, "'Calcs'!I3", "0");
        var logical = FormulaAstParser.FlattenVisible(tree).First(node => node.Info == "logical_test");

        Assert.True(logical.SourceLength > 0);
        Assert.Equal("I$3>4", formula.Substring(logical.SourceStart, logical.SourceLength));
        Assert.Equal("#7FDBFF", FormulaHighlight.CyanHex);
        var spans = FormulaHighlight.ForSelectedRows(new[]
        {
            new FormulaHighlightSpan { Start = logical.SourceStart, Length = logical.SourceLength, IsOrigin = false }
        });
        Assert.Single(spans);
        Assert.Equal(logical.SourceStart, spans[0].Start);
    }

    [Fact]
    public void FormulaHighlight_skips_origin_row()
    {
        var spans = FormulaHighlight.ForSelectedRows(new[]
        {
            new FormulaHighlightSpan { Start = 0, Length = 10, IsOrigin = true }
        });
        Assert.Empty(spans);
    }

    [Fact]
    public void AttachValidationSource_adds_validation_child()
    {
        var tree = FormulaAstParser.Parse("", "'S'!A1", "x");
        FormulaAstParser.AttachValidationSource(tree, new ValidationListSource
        {
            Label = "$A$1:$A$10",
            Location = "'S'!$A$1:$A$10",
            IsNamedRange = false
        });

        var row = FormulaAstParser.FlattenVisible(tree).Single(node => node.Info == "validation");
        Assert.Equal("$A$1:$A$10", row.Label);
        Assert.Equal("'S'!$A$1:$A$10", row.Location);
    }
}
