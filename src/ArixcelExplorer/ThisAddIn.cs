using System;
using ArixcelExplorer.ComApi;
using ArixcelExplorer.Services;

namespace ArixcelExplorer;

public partial class ThisAddIn
{
    private ArixcelComApi? _comApi;

    private void ThisAddIn_Startup(object sender, EventArgs e)
    {
        try
        {
            AddInLog.Info("ThisAddIn_Startup");
            AddInCoordinator.Initialize(Application);
            EnsureComApi().Attach(Application);
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
        }
    }

    private void ThisAddIn_Shutdown(object sender, EventArgs e)
    {
        try
        {
            AddInCoordinator.CloseAllExplorers();
            AddInCoordinator.ClearFormulaMap();
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
        }

        _comApi = null;
    }

    protected override object RequestComAddInAutomationService()
    {
        AddInLog.Info("RequestComAddInAutomationService");
        try
        {
            if (Application != null)
            {
                AddInCoordinator.Initialize(Application);
            }
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
        }

        var api = EnsureComApi();
        api.Attach(Application);
        return api;
    }

    private ArixcelComApi EnsureComApi() => _comApi ??= new ArixcelComApi();

    private void InternalStartup()
    {
        Startup += ThisAddIn_Startup;
        Shutdown += ThisAddIn_Shutdown;
    }
}
