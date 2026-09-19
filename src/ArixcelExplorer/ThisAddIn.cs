using System;
using System.Runtime.InteropServices;
using ArixcelExplorer.ComApi;
using ArixcelExplorer.Services;
using ArixcelExplorer.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer;

/// <summary>
/// Excel add-in entry point. When converted to a full VSTO project in Visual Studio,
/// this class derives from Microsoft.Office.Tools.AddInBase.
/// </summary>
[ComVisible(true)]
public sealed class ThisAddIn
{
    private Excel.Application? _application;
    private ArixcelComApi? _comApi;

    public void Startup(Excel.Application application)
    {
        _application = application;
        AddInCoordinator.Initialize(application);
        _comApi = new ArixcelComApi();
    }

    public void Shutdown()
    {
        AddInCoordinator.ClearFormulaMap();
        _application = null;
        _comApi = null;
    }

    public object CreateRibbonExtensibilityObject() => new ArixcelRibbon();

    public object? GetComApi() => _comApi;
}
