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
    public void Parse_splits_same_sheet_product_into_cell_and_comparison_rows()
    {
        const string formula = "=$D$8/4*$D$7*(E31>=$D$16)*(E31<=$D$17)";
        const string origin = "'Net Investor Returns (Insti)'!I36";
        var tree = FormulaAstParser.Parse(formula, origin, "114732.47");
        var rows = FormulaAstParser.FlattenVisible(tree);
        var labels = rows.Select(DisplayLabel).ToArray();

        Assert.Equal(
            new[]
            {
                "I36",
                "$D$8",
                "4",
                "$D$7",
                "E31>=$D$16",
                "E31",
                "$D$16",
                "E31<=$D$17",
                "E31",
                "$D$17"
            },
            labels);

        var d17 = Assert.Single(rows, node => node.Label == "$D$17");
        Assert.Equal(FormulaNodeKind.Reference, d17.Kind);
        Assert.Equal("'Net Investor Returns (Insti)'!$D$17", d17.Location);
        Assert.Equal("$D$17", formula.Substring(d17.SourceStart, d17.SourceLength));
        Assert.Equal("D17", TraceUtils.FormatExplorerLocation(d17.Location, origin));
    }

    [Fact]
    public void LooksLikeReference_requires_a_complete_address()
    {
        Assert.True(FormulaAstParser.LooksLikeReference("$D$8"));
        Assert.True(FormulaAstParser.LooksLikeReference("E31"));
        Assert.True(FormulaAstParser.LooksLikeReference("A1:A5"));
        Assert.False(FormulaAstParser.LooksLikeReference("$D$8/4*$D$7"));
        Assert.False(FormulaAstParser.LooksLikeReference("E31>=$D$16"));
    }

    [Fact]
    public void FormatExplorerLocation_shows_same_sheet_cell_without_sheet_or_dollars()
    {
        const string origin = "'Net Investor Returns (Insti)'!I36";
        Assert.Equal("D17", TraceUtils.FormatExplorerLocation("'Net Investor Returns (Insti)'!$D$17", origin));
        Assert.Equal("I36", TraceUtils.FormatExplorerLocation(origin, origin));
        Assert.Equal("'Other'!A1", TraceUtils.FormatExplorerLocation("'Other'!$A$1", origin));
        Assert.Equal("", TraceUtils.FormatExplorerLocation("E31>=$D$16", origin));
        Assert.Equal("", TraceUtils.FormatExplorerLocation("$D$8/4*$D$7", origin));
    }

    [Fact]
    public void CollapseFunctions_keeps_operator_rows_expanded()
    {
        var arithmetic = FormulaAstParser.Parse(
            "=$D$8/4*$D$7*(E31>=$D$16)*(E31<=$D$17)",
            "'S'!I36",
            "1");
        FormulaAstParser.CollapseFunctions(arithmetic);
        var arithRows = FormulaAstParser.FlattenVisible(arithmetic);
        Assert.Contains(arithRows, node => node.Label == "$D$17");

        var sumifs = FormulaAstParser.Parse("=SUMIFS(A1:A5,B1:B5,C1)", "'S'!G15", "1");
        FormulaAstParser.CollapseFunctions(sumifs);
        var functionRows = FormulaAstParser.FlattenVisible(sumifs);
        var function = Assert.Single(functionRows, node => node.Label.StartsWith("SUMIFS"));
        Assert.False(function.IsExpanded);
        Assert.DoesNotContain(functionRows, node => node.Info == "sum_range");
    }

    [Fact]
    public void QualifyAddress_does_not_prefix_operator_text()
    {
        Assert.Equal("", TraceUtils.QualifyAddress("$D$8/4*$D$7", "'S'!I36"));
        Assert.Equal("'S'!$D$8", TraceUtils.QualifyAddress("$D$8", "'S'!I36"));
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

    private static string DisplayLabel(FormulaAstNode node)
    {
        if (node.Kind == FormulaNodeKind.Root)
        {
            var parsed = TraceUtils.ParseWorksheetScopedAddress(node.Label);
            return parsed?.RangeAddress ?? node.Label;
        }

        return node.Label;
    }
}
