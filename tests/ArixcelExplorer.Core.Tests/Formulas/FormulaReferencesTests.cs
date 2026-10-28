using System.Linq;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;
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
    public void Sumifs_formula_pane_blues_only_references_and_grows_for_wrap()
    {
        const string formula =
            "=SUMIFS('Cashflow Forecast'!$C$77:$CT$77;'Cashflow Forecast'!$C$74:$CT$74;\">=\"&O$44;'Cashflow Forecast'!$C$74:$CT$74;\"<=\"&O$45)";
        var spans = FormulaTokenHits.Spans(formula, "Summary");
        var linked = string.Concat(spans.Select(span => span.Text));

        Assert.Contains("'Cashflow Forecast'!$C$77:$CT$77", linked);
        Assert.Contains("O$44", linked);
        Assert.Contains("O$45", linked);
        Assert.DoesNotContain("SUMIFS", linked);
        Assert.DoesNotContain("&", linked);
        Assert.DoesNotContain(">=", linked);
        Assert.DoesNotContain(";", linked);
        Assert.Null(FormulaTokenHits.HitTest(formula, "Summary", formula.IndexOf("SUMIFS", System.StringComparison.Ordinal)));
        var cell = formula.IndexOf("O$44", System.StringComparison.Ordinal);
        Assert.Equal("O$44", FormulaTokenHits.HitTest(formula, "Summary", cell)!.Text);

        Assert.True(FormulaPaneLayout.HeightFor(formula, 480) > ExplorerChrome.FormulaPaneHeight);
        Assert.Equal(ExplorerChrome.FormulaPaneHeight, FormulaPaneLayout.HeightFor("=D4+E4", 480));
        Assert.Equal(ExplorerChrome.FormulaPaneMaxHeight, FormulaPaneLayout.HeightFor(new string('A', 8000), 180));
        Assert.True(TraceUtils.AddressesReferToSameRange(
            "'Cashflow Forecast'!$C$77:$CT$77",
            "'Cashflow Forecast'!C77:CT77"));
        Assert.Equal("=SUMIFS(A1)", FormulaEdit.Normalize("SUMIFS(A1)"));
    }

    [Fact]
    public void Semicolon_sumifs_nests_each_criteria_pair()
    {
        const string formula =
            "=SUMIFS('Cashflow Forecast'!$C$77:$CT$77;'Cashflow Forecast'!$C$74:$CT$74;\">=\"&O$44;'Cashflow Forecast'!$C$74:$CT$74;\"<=\"&O$45)";
        var tree = FormulaAstParser.Parse(formula, "'Summary'!O46", "1");
        var sumifs = Assert.Single(tree.Children);
        Assert.Equal("SUMIFS", sumifs.Label);
        Assert.Equal("Filter 1", sumifs.Children[1].Label);
        Assert.Equal("Filter 2", sumifs.Children[2].Label);
        Assert.Equal("criteria1", sumifs.Children[1].Children[1].Info);
        Assert.Equal("\">=\"&O$44", sumifs.Children[1].Children[1].Label);
        Assert.Equal("O$45", sumifs.Children[2].Children[1].Children[1].Label);
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
