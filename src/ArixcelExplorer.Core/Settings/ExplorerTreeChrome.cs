namespace ArixcelExplorer.Core.Settings;

public static class ExplorerTreeChrome
{
    public const double IndentPerLevel = 12;
    public const double GlyphSize = 13;

    public static string Glyph(bool hasChildren, bool isExpanded) =>
        hasChildren ? (isExpanded ? "-" : "+") : "";

    public static double IndentWidth(int indent) => indent * IndentPerLevel;
}
