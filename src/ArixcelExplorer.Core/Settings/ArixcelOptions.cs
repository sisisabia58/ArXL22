using System.Collections.Generic;

namespace ArixcelExplorer.Core.Settings;

public enum ExplorerCloseBehavior
{
    EnterKeepsSelection,
    EscNavigatesBack
}

public sealed class ArixcelOptions
{
    public ExplorerCloseBehavior CloseBehavior { get; set; } = ExplorerCloseBehavior.EnterKeepsSelection;
    public int MaxDependentsBeforeWarning { get; set; } = 500;
    public int TraceMaxDepth { get; set; } = 10;
    public int TraceSafetyLimit { get; set; } = 500;
    public bool ConfirmLargeDependentScan { get; set; } = true;

    public Audit.AutoColorPalette MapPalette { get; set; } = new();

    public static ArixcelOptions Default { get; } = new();
}

public sealed class UtilityShortcutDefinition
{
    public string Id { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string Description { get; set; } = "";
    public bool Enabled { get; set; } = true;
}

public static class DefaultUtilityShortcuts
{
    public static IReadOnlyList<UtilityShortcutDefinition> All { get; } = new[]
    {
        new UtilityShortcutDefinition
        {
            Id = "unhide-all-sheets",
            DisplayName = "Unhide All Worksheets",
            Description = "Makes every hidden worksheet visible."
        },
        new UtilityShortcutDefinition
        {
            Id = "select-current-region",
            DisplayName = "Select Current Region",
            Description = "Selects the contiguous block around the active cell."
        },
        new UtilityShortcutDefinition
        {
            Id = "toggle-formula-view",
            DisplayName = "Toggle Formula View",
            Description = "Shows or hides formulas in the sheet."
        }
    };
}
