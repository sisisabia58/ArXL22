namespace ArixcelExplorer.Core.Settings;

public static class InteriorFill
{
    /// <summary>Excel xlNone. A cell with this ColorIndex has no direct fill.</summary>
    public const int NoFillColorIndex = -4142;

    public static bool IsNoFill(int colorIndex) => colorIndex == NoFillColorIndex;
}
