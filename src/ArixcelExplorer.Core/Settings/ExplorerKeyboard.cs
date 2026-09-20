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
}
