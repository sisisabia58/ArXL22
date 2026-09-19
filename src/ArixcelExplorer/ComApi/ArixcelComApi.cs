using System.Runtime.InteropServices;
using ArixcelExplorer.Services;

namespace ArixcelExplorer.ComApi;

[ComVisible(true)]
[Guid("B2C3D4E5-F6A7-8901-BCDE-F12345678901")]
[ClassInterface(ClassInterfaceType.None)]
[ProgId("ArixcelExplorer.ComApi")]
public sealed class ArixcelComApi : IArixcelComApi
{
    public void OpenExplorer() => AddInCoordinator.OpenExplorer();
    public void OpenDependents() => AddInCoordinator.OpenDependents();
    public void OpenFormulaMap() => AddInCoordinator.OpenFormulaMap();
    public void OpenCalculationFlow() => AddInCoordinator.OpenCalculationFlow();
    public void OpenCompare() => AddInCoordinator.OpenCompare();
    public void OpenOptions() => AddInCoordinator.OpenOptions();
}
