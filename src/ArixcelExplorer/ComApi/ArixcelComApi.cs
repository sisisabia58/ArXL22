using System;
using System.Runtime.InteropServices;
using ArixcelExplorer.Ribbon;
using ArixcelExplorer.Services;
using Extensibility;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.ComApi;

/// <summary>
/// COM-visible API + ribbon host for the VBA companion.
/// This add-in loads even when the VSTO customization does not, so a fresh
/// workbook still gets the EXLerate tab and Ctrl+Q.
/// </summary>
[ComVisible(true)]
[Guid("B2C3D4E5-F6A7-8901-BCDE-F12345678901")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IArixcelComApi))]
[ProgId("EXLerateExplorer.ComApi")]
public sealed class ArixcelComApi : IArixcelComApi, IDTExtensibility2, IRibbonExtensibility, ICustomQueryInterface
{
    private static readonly Guid IidIDispatch = new("00020400-0000-0000-C000-000000000046");
    private readonly ArixcelRibbon _ribbon = new();
    private Excel.Application? _excel;

    public void OpenExplorer() => Invoke(AddInCoordinator.OpenExplorer);
    public void OpenDependents() => Invoke(AddInCoordinator.OpenDependents);
    public void OpenFormulaMap() => Invoke(AddInCoordinator.OpenFormulaMap);
    public void OpenCalculationFlow() => Invoke(AddInCoordinator.OpenCalculationFlow);
    public void OpenCompare() => Invoke(AddInCoordinator.OpenCompare);
    public void OpenOptions() => Invoke(AddInCoordinator.OpenOptions);
    public void ReturnToOrigin() => Invoke(AddInCoordinator.ReturnToOrigin);
    public void CloseAllExplorers() => Invoke(AddInCoordinator.CloseAllExplorers);
    public void DispatchExplorerKey(string keyName) => Invoke(() => AddInCoordinator.DispatchExplorerKey(keyName));

    public string GetCustomUI(string ribbonId) => RibbonXml.Load("ComApi", ribbonId);
    public void OnLoad(IRibbonUI ribbonUi) => _ribbon.OnLoad(ribbonUi);
    public void OnExplorePrecedents(IRibbonControl control) => Invoke(() => _ribbon.OnExplorePrecedents(control));
    public void OnExploreDependents(IRibbonControl control) => Invoke(() => _ribbon.OnExploreDependents(control));
    public void OnOptions(IRibbonControl control) => Invoke(() => _ribbon.OnOptions(control));
    public void OnCloseAll(IRibbonControl control) => Invoke(() => _ribbon.OnCloseAll(control));
    public void OnFormulaMap(IRibbonControl control) => Invoke(() => _ribbon.OnFormulaMap(control));
    public void OnCalculationFlow(IRibbonControl control) => Invoke(() => _ribbon.OnCalculationFlow(control));
    public void OnCompare(IRibbonControl control) => Invoke(() => _ribbon.OnCompare(control));
    public void OnClearMap(IRibbonControl control) => Invoke(() => _ribbon.OnClearMap(control));
    public void OnUnhideSheets(IRibbonControl control) => Invoke(() => _ribbon.OnUnhideSheets(control));
    public void OnSelectRegion(IRibbonControl control) => Invoke(() => _ribbon.OnSelectRegion(control));
    public void OnToggleFormulas(IRibbonControl control) => Invoke(() => _ribbon.OnToggleFormulas(control));

    public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
    {
        ppv = IntPtr.Zero;
        try
        {
            if (iid == typeof(IRibbonExtensibility).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IRibbonExtensibility), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }

            if (iid == IidIDispatch)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IArixcelComApi), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
            return CustomQueryInterfaceResult.NotHandled;
        }

        return CustomQueryInterfaceResult.NotHandled;
    }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        try
        {
            AddInLog.Info($"ComApi.OnConnection mode={connectMode}");
            if (addInInst is COMAddIn comAddIn)
            {
                comAddIn.Object = this;
            }

            Attach(application as Excel.Application);
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
        }
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        AddInLog.Info($"ComApi.OnDisconnection mode={removeMode}");
    }

    public void OnAddInsUpdate(ref Array custom) { }

    public void OnStartupComplete(ref Array custom)
    {
        try
        {
            AddInLog.Info("ComApi.OnStartupComplete");
            Attach(_excel);
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
        }
    }

    public void OnBeginShutdown(ref Array custom) { }

    internal void Attach(Excel.Application? application)
    {
        if (application != null)
        {
            _excel = application;
        }

        TryInitialize(_excel);
    }

    private void TryInitialize(Excel.Application? application)
    {
        if (application != null)
        {
            AddInCoordinator.Initialize(application);
            return;
        }

        try
        {
            var excel = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            _excel = excel;
            AddInCoordinator.Initialize(excel);
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
        }
    }

    private void Invoke(Action action)
    {
        try
        {
            Attach(_excel);
            action();
        }
        catch (Exception ex)
        {
            throw AddInCoordinator.Unwrap(ex);
        }
    }
}
