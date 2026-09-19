using System;
using System.Runtime.InteropServices;
using ArixcelExplorer.Services;
using Extensibility;

namespace ArixcelExplorer.ComApi;

/// <summary>
/// COM-visible API surface for the VBA companion add-in (Ctrl+Q / Ctrl+Shift+Q).
/// Implements IDTExtensibility2 so Excel lists it in COMAddIns.
/// </summary>
[ComVisible(true)]
[Guid("B2C3D4E5-F6A7-8901-BCDE-F12345678901")]
[ClassInterface(ClassInterfaceType.None)]
[ProgId("ArixcelExplorer.ComApi")]
public sealed class ArixcelComApi : IArixcelComApi, IDTExtensibility2
{
    public void OpenExplorer() => AddInCoordinator.OpenExplorer();
    public void OpenDependents() => AddInCoordinator.OpenDependents();
    public void OpenFormulaMap() => AddInCoordinator.OpenFormulaMap();
    public void OpenCalculationFlow() => AddInCoordinator.OpenCalculationFlow();
    public void OpenCompare() => AddInCoordinator.OpenCompare();
    public void OpenOptions() => AddInCoordinator.OpenOptions();
    public void ReturnToOrigin() => AddInCoordinator.ReturnToOrigin();
    public void CloseAllExplorers() => AddInCoordinator.CloseAllExplorers();

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        if (addInInst is Microsoft.Office.Core.COMAddIn comAddIn)
        {
            comAddIn.Object = this;
        }
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) { }

    public void OnAddInsUpdate(ref Array custom) { }

    public void OnStartupComplete(ref Array custom) { }

    public void OnBeginShutdown(ref Array custom) { }
}
