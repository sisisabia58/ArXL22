using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ArixcelExplorer.Services;
using Office = Microsoft.Office.Core;

namespace ArixcelExplorer.Ribbon;

[ComVisible(true)]
public sealed class ArixcelRibbon : Office.IRibbonExtensibility
{
    private Office.IRibbonUI? _ribbon;

    public string GetCustomUI(string ribbonId)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = "ArixcelExplorer.Ribbon.ArixcelRibbon.xml";
        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream != null)
        {
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        var path = Path.Combine(Path.GetDirectoryName(assembly.Location) ?? "", "Ribbon", "ArixcelRibbon.xml");
        return File.Exists(path) ? File.ReadAllText(path) : "<customUI/>";
    }

    public void OnLoad(Office.IRibbonUI ribbonUi) => _ribbon = ribbonUi;

    public void OnExplorePrecedents(Office.IRibbonControl control) => AddInCoordinator.OpenExplorer();
    public void OnExploreDependents(Office.IRibbonControl control) => AddInCoordinator.OpenDependents();
    public void OnOptions(Office.IRibbonControl control) => AddInCoordinator.OpenOptions();
    public void OnFormulaMap(Office.IRibbonControl control) => AddInCoordinator.OpenFormulaMap();
    public void OnCalculationFlow(Office.IRibbonControl control) => AddInCoordinator.OpenCalculationFlow();
    public void OnCompare(Office.IRibbonControl control) => AddInCoordinator.OpenCompare();
    public void OnClearMap(Office.IRibbonControl control) => AddInCoordinator.ClearFormulaMap();
    public void OnUnhideSheets(Office.IRibbonControl control) =>
        AddInCoordinator.ExecuteUtilityShortcut("unhide-all-sheets");
    public void OnSelectRegion(Office.IRibbonControl control) =>
        AddInCoordinator.ExecuteUtilityShortcut("select-current-region");
    public void OnToggleFormulas(Office.IRibbonControl control) =>
        AddInCoordinator.ExecuteUtilityShortcut("toggle-formula-view");
}
