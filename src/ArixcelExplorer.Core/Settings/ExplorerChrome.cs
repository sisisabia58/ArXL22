namespace ArixcelExplorer.Core.Settings;

/// <summary>
/// Explorer window chrome aligned to Arixcel Explorer 9 (side-by-side on =D4+E4):
/// a wider tree, 12pt text, and a formula row tall enough for underlined links.
/// </summary>
public static class ExplorerChrome
{
    public const double FontSize = 12;

    public const double ExplorerWidth = 560;
    public const double ExplorerHeight = 420;
    public const double ExplorerMinWidth = 520;
    public const double ExplorerMinHeight = 300;

    public const double LegacyCompactWidth = 428;
    public const double LegacyCompactHeight = 338;

    public const double DependentsWidth = 520;
    public const double DependentsHeight = 340;
    public const double DependentsMinWidth = 400;
    public const double DependentsMinHeight = 240;

    public const double OptionsWidth = 380;
    public const double OptionsHeight = 292;

    public const double RowHeight = 20;
    public const double HeaderHeight = 24;
    public const double FormulaPaneHeight = 32;

    public static bool IsLegacyCompact(double width, double height) =>
        System.Math.Abs(width - LegacyCompactWidth) < 2 &&
        System.Math.Abs(height - LegacyCompactHeight) < 2;

    public const double ElementWidth = 190;
    public const double InfoWidth = 88;
    public const double ValueWidth = 140;
    public const double LocationWidth = 110;
    public const double CountWidth = 48;

    public const double OkButtonWidth = 75;
    public const double OkButtonHeight = 23;
    public const double MoreButtonWidth = 78;

    public const double ExcelOffsetX = 48;
    public const double ExcelOffsetY = 88;
}
