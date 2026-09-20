using System;
using System.Runtime.InteropServices;
using ArixcelExplorer.Services;
using Microsoft.Office.Core;

namespace ArixcelExplorer.Ribbon;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class ArixcelRibbon : IRibbonExtensibility
{
    private static IRibbonUI? SharedUi;
    private IRibbonUI? _ribbon;

    public string GetCustomUI(string ribbonId) => RibbonXml.Load("VSTO", ribbonId);

    public void OnLoad(IRibbonUI ribbonUi)
    {
        _ribbon = ribbonUi;
        SharedUi = ribbonUi;
        AddInLog.Info("Ribbon OnLoad");
    }

    public static void Invalidate()
    {
        try
        {
            SharedUi?.Invalidate();
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
        }
    }

    public void OnExplorePrecedents(IRibbonControl control) => AddInCoordinator.OpenExplorer();
    public void OnExploreDependents(IRibbonControl control) => AddInCoordinator.OpenDependents();
    public void OnOptions(IRibbonControl control) => AddInCoordinator.OpenOptions();
    public void OnCloseAll(IRibbonControl control) => AddInCoordinator.CloseAllExplorers();
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
