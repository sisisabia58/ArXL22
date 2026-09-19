using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C0396-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IRibbonExtensibility
{
    string GetCustomUI(string ribbonId);
}

[ComImport]
[Guid("000C0395-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IRibbonUI
{
    void Invalidate();
    void InvalidateControl(string bstrControlId);
    void InvalidateControlMso(string bstrControlIdMso);
    void ActivateTab(string bstrControlId);
    void ActivateTabMso(string bstrControlIdMso);
    void ActivateTabQ(string bstrControlId, string bstrQ);
}

[ComImport]
[Guid("C39E5812-4692-44F2-B601-8931F1192A74")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IRibbonControl
{
    string Id { get; }
    object Context { get; }
    string Tag { get; }
}
