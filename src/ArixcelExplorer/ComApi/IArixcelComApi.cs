using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace ArixcelExplorer.ComApi;

[ComVisible(true)]
[Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IArixcelComApi
{
    void OpenExplorer();
    void OpenDependents();
    void OpenFormulaMap();
    void OpenCalculationFlow();
    void OpenCompare();
    void OpenOptions();
    void ReturnToOrigin();
    void CloseAllExplorers();
    void DispatchExplorerKey(string keyName);

    string GetCustomUI(string ribbonId);
    void OnLoad(IRibbonUI ribbonUi);
    void OnExplorePrecedents(IRibbonControl control);
    void OnExploreDependents(IRibbonControl control);
    void OnOptions(IRibbonControl control);
    void OnCloseAll(IRibbonControl control);
    void OnFormulaMap(IRibbonControl control);
    void OnCalculationFlow(IRibbonControl control);
    void OnCompare(IRibbonControl control);
    void OnClearMap(IRibbonControl control);
    void OnUnhideSheets(IRibbonControl control);
    void OnSelectRegion(IRibbonControl control);
    void OnToggleFormulas(IRibbonControl control);
}
