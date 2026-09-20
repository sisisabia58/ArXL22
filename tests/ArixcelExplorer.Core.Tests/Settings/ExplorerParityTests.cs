using System;
using System.IO;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;
using Xunit;

namespace ArixcelExplorer.Core.Tests.Settings;

public sealed class ExplorerParityTests
{
    [Fact]
    public void ExplorerChrome_uses_compact_original_sizes()
    {
        Assert.True(ExplorerChrome.ExplorerWidth < 480);
        Assert.True(ExplorerChrome.ExplorerHeight < 400);
        Assert.True(ExplorerChrome.FormulaPaneHeight <= 28);
        Assert.Equal(18, ExplorerChrome.RowHeight);
    }

    [Theory]
    [InlineData(0, ExplorerCtrlQAction.CycleExpand)]
    [InlineData(-1, ExplorerCtrlQAction.CycleExpand)]
    [InlineData(1, ExplorerCtrlQAction.Drill)]
    [InlineData(4, ExplorerCtrlQAction.Drill)]
    public void ResolveCtrlQ_cycles_on_origin_and_drills_otherwise(int selectedIndex, ExplorerCtrlQAction expected)
    {
        Assert.Equal(expected, ExplorerKeyboard.ResolveCtrlQ(selectedIndex));
    }

    [Fact]
    public void ShouldDrillDependents_only_when_not_origin()
    {
        Assert.False(ExplorerKeyboard.ShouldDrillDependents(0));
        Assert.False(ExplorerKeyboard.ShouldDrillDependents(-1));
        Assert.True(ExplorerKeyboard.ShouldDrillDependents(2));
    }

    [Fact]
    public void FormulaTokenHits_finds_reference_at_character_index()
    {
        var formula = "=B1+Sheet2!C3";
        var b1 = FormulaTokenHits.HitTest(formula, "Sheet1", 1);
        Assert.NotNull(b1);
        Assert.Equal("B1", b1!.Text);
        Assert.Equal("'Sheet1'!B1", b1.Address);

        var c3 = FormulaTokenHits.HitTest(formula, "Sheet1", formula.IndexOf("C3", StringComparison.Ordinal));
        Assert.NotNull(c3);
        Assert.Equal("Sheet2!C3", c3!.Text);
        Assert.Contains("C3", c3.Address, StringComparison.OrdinalIgnoreCase);

        Assert.Null(FormulaTokenHits.HitTest(formula, "Sheet1", 0));
    }

    [Fact]
    public void HandleCtrlQ_contract_cycles_on_origin_and_drills_on_other_rows()
    {
        var cycled = 0;
        var drilled = 0;

        void Handle(int selectedIndex)
        {
            if (ExplorerKeyboard.ResolveCtrlQ(selectedIndex) == ExplorerCtrlQAction.CycleExpand)
            {
                cycled += 1;
            }
            else
            {
                drilled += 1;
            }
        }

        Handle(0);
        Handle(3);
        Assert.Equal(1, cycled);
        Assert.Equal(1, drilled);
        Assert.False(ExplorerKeyboard.ShouldDrillDependents(0));
        Assert.True(ExplorerKeyboard.ShouldDrillDependents(1));
    }

    [Fact]
    public void FormulaTokenHits_address_matches_qualified_tree_location()
    {
        var hit = FormulaTokenHits.HitTest("=B1+Sheet2!C3", "Sheet1", 1);
        var location = TraceUtils.QualifyAddress("B1", "'Sheet1'!C1");
        Assert.True(TraceUtils.AddressesReferToSameRange(hit!.Address, location));
    }

    [Fact]
    public void OptionsStore_round_trips_colors_and_window_bounds()
    {
        var path = Path.Combine(Path.GetTempPath(), "exlerate-options-test-" + Guid.NewGuid().ToString("N") + ".json");
        try
        {
            var options = new ArixcelOptions();
            options.OriginHighlight = "#AABBCC";
            options.PrecedentHighlight = "#112233";
            options.DependentHighlight = "#445566";
            options.ConfirmLargeDependentScan = false;
            options.CloseBehavior = ExplorerCloseBehavior.EscNavigatesBack;
            options.ExplorerWindow = new WindowPlacement { Left = 10, Top = 20, Width = 430, Height = 340 };
            options.DependentsWindow = new WindowPlacement { Left = 40, Top = 50, Width = 428, Height = 300 };

            OptionsStore.Save(options, path);
            var loaded = OptionsStore.Load(path);

            Assert.Equal("#AABBCC", loaded.OriginHighlight);
            Assert.Equal("#112233", loaded.PrecedentHighlight);
            Assert.Equal("#445566", loaded.DependentHighlight);
            Assert.False(loaded.ConfirmLargeDependentScan);
            Assert.Equal(ExplorerCloseBehavior.EscNavigatesBack, loaded.CloseBehavior);
            Assert.Equal(10, loaded.ExplorerWindow.Left);
            Assert.Equal(20, loaded.ExplorerWindow.Top);
            Assert.Equal(430, loaded.ExplorerWindow.Width);
            Assert.Equal(340, loaded.ExplorerWindow.Height);
            Assert.Equal(40, loaded.DependentsWindow.Left);
            Assert.True(loaded.ExplorerWindow.HasPosition);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void AddressesReferToSameRange_ignores_sheet_quoting()
    {
        Assert.True(TraceUtils.AddressesReferToSameRange("'Sheet1'!B1", "Sheet1!B1"));
        Assert.False(TraceUtils.AddressesReferToSameRange("'Sheet1'!B1", "'Sheet1'!C1"));
    }
}
