using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ArixcelExplorer.Services;
using Microsoft.Office.Core;

namespace ArixcelExplorer.Ribbon;

[ComVisible(true)]
public sealed class ArixcelRibbon : IRibbonExtensibility
{
    private IRibbonUI? _ribbon;

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

    public void OnLoad(IRibbonUI ribbonUi) => _ribbon = ribbonUi;

    public void OnExplorePrecedents(IRibbonControl control) => AddInCoordinator.OpenExplorer();
    public void OnExploreDependents(IRibbonControl control) => AddInCoordinator.OpenDependents();
    public void OnOptions(IRibbonControl control) => AddInCoordinator.OpenOptions();
    public void OnFormulaMap(IRibbonControl control) => AddInCoordinator.OpenFormulaMap();
    public void OnCalculationFlow(IRibbonControl control) => AddInCoordinator.OpenCalculationFlow();
    public void OnCompare(IRibbonControl control) => AddInCoordinator.OpenCompare();
    public void OnClearMap(IRibbonControl control) => AddInCoordinator.ClearFormulaMap();
    public void OnUnhideSheets(IRibbonControl control) =>
        AddInCoordinator.ExecuteUtilityShortcut("unhide-all-sheets");
    public void OnSelectRegion(IRibbonControl control) =>
        AddInCoordinator.ExecuteUtilityShortcut("select-current-region");
    public void OnToggleFormulas(IRibbonControl control) =>
        AddInCoordinator.ExecuteUtilityShortcut("toggle-formula-view");
}
