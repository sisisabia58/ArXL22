using System;

namespace ArixcelExplorer.Core.Settings;

/// <summary>
/// Height of the Explorer formula pane. Long formulas wrap onto extra lines
/// instead of being clipped by the one-line chrome.
/// </summary>
public static class FormulaPaneLayout
{
    public static double HeightFor(string? text, double availableWidth, double fontSize = ExplorerChrome.FontSize)
    {
        var content = string.IsNullOrEmpty(text) ? " " : text!;
        var width = availableWidth < 48 ? 48 : availableWidth;
        // Slightly narrow so a wrapped SUMIFS gains a line of breathing room rather than clipping.
        var charWidth = Math.Max(1, fontSize * 0.48);
        var charsPerLine = Math.Max(8, (int)(width / charWidth));
        var lines = 1;
        var column = 0;
        foreach (var ch in content)
        {
            if (ch == '\r') continue;
            if (ch == '\n')
            {
                lines += 1;
                column = 0;
                continue;
            }

            column += 1;
            if (column >= charsPerLine)
            {
                lines += 1;
                column = 0;
            }
        }

        var lineHeight = fontSize + 6;
        var raw = lines * lineHeight + 14;
        if (raw < ExplorerChrome.FormulaPaneHeight) return ExplorerChrome.FormulaPaneHeight;
        if (raw > ExplorerChrome.FormulaPaneMaxHeight) return ExplorerChrome.FormulaPaneMaxHeight;
        return raw;
    }
}
