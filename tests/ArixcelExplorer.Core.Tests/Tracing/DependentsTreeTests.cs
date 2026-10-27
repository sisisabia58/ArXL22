using System.Collections.Generic;
using System.Linq;
using ArixcelExplorer.Core.Tracing;
using Xunit;

namespace ArixcelExplorer.Core.Tests.Tracing;

public sealed class DependentsTreeTests
{
    private const string Origin = "'Investment Cashflows'!G44";

    [Fact]
    public void Build_roots_the_source_and_counts_coalesced_entries()
    {
        var tree = DependentsTree.Build(Origin, "46.660", SampleEntries(), groupBySheet: false);

        Assert.Equal("G44", tree.Display);
        Assert.True(tree.IsRoot);
        Assert.Equal(24, tree.Count);
        Assert.Equal(24, tree.Children.Count);
        Assert.DoesNotContain(tree.Children, node => node.Display == "G44" || node.Address.EndsWith("!G44"));
        Assert.Equal("I11", tree.Children[0].Display);
        Assert.Equal("'Investment Cashflows'!I11", tree.Children[0].Address);
        Assert.Equal("G48:G53", tree.Children.Single(node => node.Display == "G48:G53").Display);
        Assert.Equal("'Summary (Instit)'!D11:J11", tree.Children.Single(node => node.Display.Contains("Summary")).Display);
        Assert.Equal("'Net Investor Returns (Instit)'!I33", tree.Children.Single(node => node.Display.Contains("Net Investor")).Display);

        var displays = tree.Children.Select(node => node.Display).ToArray();
        Assert.Equal(new[] { "G97", "G98", "G99:G100" }, displays.SkipWhile(item => item != "G97").Take(3));
        Assert.Equal(new[] { "G104", "G105", "G106:G107" }, displays.SkipWhile(item => item != "G104").Take(3));
        Assert.Contains("I13", displays);
        Assert.Contains("I16", displays);
        Assert.Contains("H44", displays);
        Assert.Contains("G82", displays);
        Assert.Contains("G172", displays);
        Assert.DoesNotContain("G97:G100", displays);
        Assert.DoesNotContain("G104:G107", displays);
    }

    [Fact]
    public void Combine_prefers_arrow_splits_over_a_geometric_block_and_appends_other_sheets()
    {
        var combined = DependentsTree.Combine(
            Origin,
            new[]
            {
                "G97",
                "G98",
                "$G$99:$G$100",
                "'Summary (Instit)'!$D$11:$J$11",
                "'Net Investor Returns (Instit)'!I33"
            },
            new[] { "G97:G100", "G82", "G44" });

        var displays = combined.Select(entry => DependentsTree.Build(Origin, "", new[] { entry }, false).Children[0].Display).ToArray();
        Assert.Equal(
            new[]
            {
                "G97",
                "G98",
                "G99:G100",
                "G82",
                "'Summary (Instit)'!D11:J11",
                "'Net Investor Returns (Instit)'!I33"
            },
            displays);
    }

    [Fact]
    public void Combine_does_not_merge_separate_cells_when_excel_reports_them_apart()
    {
        var combined = DependentsTree.Combine(
            Origin,
            null,
            new[] { "G97", "G98", "G99", "G100", "G48:G53" });

        Assert.Equal(5, combined.Count);
        Assert.Equal("'Investment Cashflows'!G48:G53", combined[4].Address);
        Assert.DoesNotContain(combined, entry => entry.Address.Contains("G97:G100"));
    }

    [Fact]
    public void I32_row_spans_replace_partial_cells_and_keep_the_sheet_name_when_grouped()
    {
        const string origin = "'Summary (Instit)'!I32";
        var cells = new List<string>();
        foreach (var column in new[] { "J", "I", "H", "G" })
        {
            foreach (var row in new[] { 60, 61, 62, 63 })
            {
                cells.Add(column + row);
            }
        }

        var combined = DependentsTree.Combine(
            origin,
            cells,
            new[] { "D60:J60", "D61:J61", "D62:J62", "D63:J63", "D67:J67" });
        var grouped = DependentsTree.Build(origin, "", combined, groupBySheet: true);

        Assert.Equal("I32", grouped.Display);
        Assert.Equal(5, grouped.Count);
        Assert.Equal("Summary (Instit)", grouped.Children[0].Display);
        Assert.Equal(5, grouped.Children[0].Count);
        Assert.Equal(
            new[]
            {
                "'Summary (Instit)'!D60:J60",
                "'Summary (Instit)'!D61:J61",
                "'Summary (Instit)'!D62:J62",
                "'Summary (Instit)'!D63:J63",
                "'Summary (Instit)'!D67:J67"
            },
            grouped.Children[0].Children.Select(node => node.Display).ToArray());

        var flat = DependentsTree.Build(origin, "", combined, groupBySheet: false);
        Assert.Equal("D60:J60", flat.Children[0].Display);
        Assert.Equal("'Summary (Instit)'!D60:J60", flat.Children[0].Address);
    }

    [Fact]
    public void Adjacent_cells_on_one_row_coalesce_without_merging_separate_column_targets()
    {
        var entries = new List<DependentEntry>();
        foreach (var column in new[] { "D", "E", "F", "G", "H", "I", "J" })
        {
            entries.Add(new DependentEntry { Address = column + "60" });
            entries.Add(new DependentEntry { Address = column + "67" });
        }

        entries.Add(new DependentEntry { Address = "G97" });
        entries.Add(new DependentEntry { Address = "G98" });
        entries.Add(new DependentEntry { Address = "G99" });
        entries.Add(new DependentEntry { Address = "G100" });

        var tree = DependentsTree.Build(Origin, "", entries, groupBySheet: false);
        var displays = tree.Children.Select(node => node.Display).ToArray();
        Assert.Contains("D60:J60", displays);
        Assert.Contains("D67:J67", displays);
        Assert.Contains("G97", displays);
        Assert.Contains("G98", displays);
        Assert.Contains("G99", displays);
        Assert.Contains("G100", displays);
        Assert.DoesNotContain(displays, item => item == "G97:G100" || item == "G97:G98" || item == "G99:G100");
        Assert.Equal(6, tree.Count);
    }

    [Fact]
    public void Build_groups_by_sheet_without_changing_the_root_count()
    {
        var tree = DependentsTree.Build(Origin, "", SampleEntries(), groupBySheet: true);

        Assert.Equal(24, tree.Count);
        Assert.Equal(3, tree.Children.Count);
        Assert.Equal("Investment Cashflows", tree.Children[0].Display);
        Assert.Equal(22, tree.Children[0].Count);
        Assert.Equal("Summary (Instit)", tree.Children[1].Display);
        Assert.Equal("'Summary (Instit)'!D11:J11", tree.Children[1].Children[0].Display);
        Assert.Equal("'Summary (Instit)'!D11:J11", tree.Children[1].Children[0].Address);
        Assert.Equal("Net Investor Returns (Instit)", tree.Children[2].Display);
        Assert.Equal("'Net Investor Returns (Instit)'!I33", tree.Children[2].Children[0].Display);
        Assert.Equal(1, tree.Children[2].Count);
        Assert.Equal("'Investment Cashflows'!I11", tree.Children[0].Children[0].Display);
    }

    [Fact]
    public void Build_empty_dependents_is_a_root_with_count_zero()
    {
        var tree = DependentsTree.Build(Origin, "46.660", new DependentEntry[0], groupBySheet: false);

        Assert.Equal("G44", tree.Display);
        Assert.Equal(0, tree.Count);
        Assert.Empty(tree.Children);
    }

    private static IReadOnlyList<DependentEntry> SampleEntries()
    {
        var addresses = new[]
        {
            "I11", "I13", "I16", "G43", "H44", "G45", "G48:G53", "G56", "G74",
            "G75:G78", "G82", "C92", "G97", "G98", "G99:G100", "G104", "G105",
            "G106:G107", "C112", "C119", "G170", "G172",
            "'Summary (Instit)'!D11:J11",
            "'Net Investor Returns (Instit)'!I33",
            "G44"
        };
        var entries = new List<DependentEntry>(addresses.Length);
        foreach (var address in addresses)
        {
            entries.Add(new DependentEntry { Address = address, Count = 1 });
        }

        return entries;
    }
}
