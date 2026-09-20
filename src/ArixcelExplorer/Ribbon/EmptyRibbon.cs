using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace ArixcelExplorer.Ribbon;

[ComVisible(true)]
public sealed class EmptyRibbon : IRibbonExtensibility
{
    public string GetCustomUI(string ribbonId)
    {
        _ = ribbonId;
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" />
            """;
    }
}
