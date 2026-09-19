using System;
using ArixcelExplorer.ComApi;
using ArixcelExplorer.Services;

namespace ArixcelExplorer;

public partial class ThisAddIn
{
    private ArixcelComApi? _comApi;

    private void ThisAddIn_Startup(object sender, EventArgs e)
    {
        AddInCoordinator.Initialize(Application);
        _comApi = new ArixcelComApi();
    }

    private void ThisAddIn_Shutdown(object sender, EventArgs e)
    {
        AddInCoordinator.CloseAllExplorers();
        AddInCoordinator.ClearFormulaMap();
        _comApi = null;
    }

    #region VSTO generated code

    private void InternalStartup()
    {
        Startup += new EventHandler(ThisAddIn_Startup);
        Shutdown += new EventHandler(ThisAddIn_Shutdown);
    }

    #endregion
}
