using System.Globalization;

namespace ArixcelExplorer.Core.Settings;

public static class ExcelOleColor
{
    public static bool IsValidHex(string? hex)
    {
        var body = NormalizeBody(hex);
        return body != null &&
               int.TryParse(body, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _);
    }

    public static int FromHex(string? hex, int fallback = 0xFFFFFF)
    {
        var body = NormalizeBody(hex);
        if (body == null) return fallback;
        if (!int.TryParse(body, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
        {
            return fallback;
        }

        var r = (rgb >> 16) & 0xFF;
        var g = (rgb >> 8) & 0xFF;
        var b = rgb & 0xFF;
        return r | (g << 8) | (b << 16);
    }

    private static string? NormalizeBody(string? hex)
    {
        if (hex == null) return null;
        var text = hex.Trim();
        if (text.Length == 0) return null;
        if (text.StartsWith("#", System.StringComparison.Ordinal))
        {
            text = text.Substring(1);
        }

        return text.Length == 6 ? text : null;
    }
}
