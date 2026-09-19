using System;
using ArixcelExplorer.Core.Settings;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.Services;

public sealed class UtilityShortcutService
{
    private readonly Excel.Application _app;

    public UtilityShortcutService(Excel.Application application)
    {
        _app = application;
    }

    public void Execute(string shortcutId)
    {
        switch (shortcutId)
        {
            case "unhide-all-sheets":
                UnhideAllSheets();
                break;
            case "select-current-region":
                (_app.Selection as Excel.Range)?.CurrentRegion?.Select();
                break;
            case "toggle-formula-view":
                _app.DisplayFormulas = !_app.DisplayFormulas;
                break;
            default:
                throw new InvalidOperationException($"Unknown shortcut: {shortcutId}");
        }
    }

    private void UnhideAllSheets()
    {
        foreach (Excel.Worksheet sheet in _app.Worksheets)
        {
            sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
        }
    }
}
