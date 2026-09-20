using System;

namespace ArixcelExplorer.Core.Settings;

public enum ExplorerCtrlQAction
{
    CycleExpand,
    Drill
}

public static class ExplorerKeyboard
{
    public static ExplorerCtrlQAction ResolveCtrlQ(int selectedIndex) =>
        selectedIndex <= 0 ? ExplorerCtrlQAction.CycleExpand : ExplorerCtrlQAction.Drill;

    public static bool ShouldDrillDependents(int selectedIndex) => selectedIndex > 0;

    public static string? CanonicalKeyName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;

        var text = name!.Trim().Trim('{', '}').Trim();
        if (text.Equals("UP", StringComparison.OrdinalIgnoreCase)) return "Up";
        if (text.Equals("DOWN", StringComparison.OrdinalIgnoreCase)) return "Down";
        if (text.Equals("LEFT", StringComparison.OrdinalIgnoreCase)) return "Left";
        if (text.Equals("RIGHT", StringComparison.OrdinalIgnoreCase)) return "Right";
        if (text.Equals("RETURN", StringComparison.OrdinalIgnoreCase) ||
            text.Equals("ENTER", StringComparison.OrdinalIgnoreCase))
        {
            return "Enter";
        }

        if (text.Equals("ESC", StringComparison.OrdinalIgnoreCase) ||
            text.Equals("ESCAPE", StringComparison.OrdinalIgnoreCase))
        {
            return "Escape";
        }

        return null;
    }
}
