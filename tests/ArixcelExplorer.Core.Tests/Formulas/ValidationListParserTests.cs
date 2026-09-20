using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Settings;
using Xunit;

namespace ArixcelExplorer.Core.Tests.Formulas;

public sealed class ValidationListParserTests
{
    [Theory]
    [InlineData("=$A$1:$A$10", "$A$1:$A$10", false)]
    [InlineData("=Sheet2!$B$1:$B$20", "Sheet2!$B$1:$B$20", false)]
    [InlineData("=DeptList", "DeptList", true)]
    [InlineData("='Sheet 1'!$A$1:$A$10", "'Sheet 1'!$A$1:$A$10", false)]
    [InlineData("\"=$A$1:$A$10\"", "$A$1:$A$10", false)]
    [InlineData("=R1C1:R10C1", "R1C1:R10C1", false)]
    [InlineData("='Investment Cashflows'!R48C3:R48C32", "'Investment Cashflows'!R48C3:R48C32", false)]
    public void TryParse_accepts_range_and_named_list_sources(string formula1, string expectedLocation, bool named)
    {
        var parsed = ValidationListParser.TryParse(formula1, "Sheet1");
        Assert.NotNull(parsed);
        Assert.Equal(expectedLocation, parsed!.Location);
        Assert.Equal(named, parsed.IsNamedRange);
    }

    [Theory]
    [InlineData("Yes,No")]
    [InlineData("=Yes,No")]
    [InlineData("")]
    [InlineData(null)]
    public void TryParse_skips_literal_comma_lists(string? formula1)
    {
        Assert.Null(ValidationListParser.TryParse(formula1, "Sheet1"));
    }

    [Fact]
    public void IsR1C1_detects_absolute_r1c1_ranges()
    {
        Assert.True(ValidationListParser.IsR1C1("R1C1:R10C1"));
        Assert.True(ValidationListParser.IsR1C1("'Sheet'!R48C3:R48C32"));
        Assert.False(ValidationListParser.IsR1C1("$A$1:$A$10"));
    }
}

public sealed class ExplorerKeyboardTests
{
    [Theory]
    [InlineData("Up", "Up")]
    [InlineData("{DOWN}", "Down")]
    [InlineData("left", "Left")]
    [InlineData("RIGHT", "Right")]
    [InlineData("RETURN", "Enter")]
    [InlineData("{ESC}", "Escape")]
    public void CanonicalKeyName_maps_arrow_aliases(string raw, string expected)
    {
        Assert.Equal(expected, ExplorerKeyboard.CanonicalKeyName(raw));
    }
}
