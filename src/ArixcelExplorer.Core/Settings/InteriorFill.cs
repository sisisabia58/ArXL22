namespace ArixcelExplorer.Core.Settings;

public static class InteriorFill
{
    /// <summary>Excel xlNone. Used for both ColorIndex and Pattern when a cell has no direct fill.</summary>
    public const int NoFillColorIndex = -4142;

    /// <summary>
    /// An RGB fill such as the yellow input cells sets Pattern to xlSolid (1) and often
    /// leaves ColorIndex at xlNone. That is still a fill. Only Pattern xlNone is empty.
    /// </summary>
    public static bool IsNoFill(int colorIndex, int? pattern)
    {
        if (pattern is int value) return value == NoFillColorIndex;
        return colorIndex == NoFillColorIndex;
    }
}
