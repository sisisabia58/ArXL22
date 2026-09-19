using System.Runtime.InteropServices;

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
}
