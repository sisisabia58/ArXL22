namespace ArixcelExplorer.Core.Settings;

/// <summary>
/// Compact sizes from original Arixcel Explorer (Win32 list dialog) as documented
/// in product screenshots (Claritix / arixcel.com). Live Excel was not open during
/// Wave 0 measurement; first-open outer size is ~428x338 with ~18px rows.
/// </summary>
public static class ExplorerChrome
{
    public const double ExplorerWidth = 428;
    public const double ExplorerHeight = 338;
    public const double ExplorerMinWidth = 360;
    public const double ExplorerMinHeight = 240;

    public const double DependentsWidth = 428;
    public const double DependentsHeight = 300;
    public const double DependentsMinWidth = 360;
    public const double DependentsMinHeight = 220;

    public const double OptionsWidth = 380;
    public const double OptionsHeight = 292;

    public const double RowHeight = 18;
    public const double HeaderHeight = 22;
    public const double FormulaPaneHeight = 24;

    public const double ElementWidth = 150;
    public const double InfoWidth = 72;
    public const double ValueWidth = 72;
    public const double LocationWidth = 110;
    public const double CountWidth = 48;

    public const double OkButtonWidth = 75;
    public const double OkButtonHeight = 23;

    public const double ExcelOffsetX = 48;
    public const double ExcelOffsetY = 88;
}
