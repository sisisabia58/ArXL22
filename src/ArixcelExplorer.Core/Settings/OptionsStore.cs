using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ArixcelExplorer.Core.Settings;

public static class OptionsStore
{
    public static string DefaultPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "EXLerate",
        "options.json");

    public static void Save(ArixcelOptions options, string? path = null)
    {
        var file = path ?? DefaultPath;
        var directory = Path.GetDirectoryName(file);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(file, Serialize(options ?? new ArixcelOptions()));
    }

    public static ArixcelOptions Load(string? path = null)
    {
        var file = path ?? DefaultPath;
        if (!File.Exists(file)) return new ArixcelOptions();

        try
        {
            return Deserialize(File.ReadAllText(file));
        }
        catch
        {
            return new ArixcelOptions();
        }
    }

    public static ArixcelOptions Clone(ArixcelOptions options) =>
        Deserialize(Serialize(options ?? new ArixcelOptions()));

    public static string Serialize(ArixcelOptions options)
    {
        var builder = new StringBuilder();
        builder.AppendLine("{");
        WriteString(builder, "closeBehavior", options.CloseBehavior.ToString(), comma: true);
        WriteString(builder, "originHighlight", options.OriginHighlight, comma: true);
        WriteString(builder, "precedentHighlight", options.PrecedentHighlight, comma: true);
        WriteString(builder, "dependentHighlight", options.DependentHighlight, comma: true);
        WriteInt(builder, "maxDependentsBeforeWarning", options.MaxDependentsBeforeWarning, comma: true);
        WriteInt(builder, "traceMaxDepth", options.TraceMaxDepth, comma: true);
        WriteInt(builder, "traceSafetyLimit", options.TraceSafetyLimit, comma: true);
        WriteBool(builder, "confirmLargeDependentScan", options.ConfirmLargeDependentScan, comma: true);
        WritePlacement(builder, "explorerWindow", options.ExplorerWindow, comma: true);
        WritePlacement(builder, "dependentsWindow", options.DependentsWindow, comma: false);
        builder.AppendLine();
        builder.Append('}');
        return builder.ToString();
    }

    public static ArixcelOptions Deserialize(string json)
    {
        var options = new ArixcelOptions();
        if (string.IsNullOrWhiteSpace(json)) return options;

        var close = ReadString(json, "closeBehavior");
        if (Enum.TryParse(close, ignoreCase: true, out ExplorerCloseBehavior behavior))
        {
            options.CloseBehavior = behavior;
        }

        options.OriginHighlight = ReadString(json, "originHighlight") ?? options.OriginHighlight;
        options.PrecedentHighlight = ReadString(json, "precedentHighlight") ?? options.PrecedentHighlight;
        options.DependentHighlight = ReadString(json, "dependentHighlight") ?? options.DependentHighlight;
        options.MaxDependentsBeforeWarning = ReadInt(json, "maxDependentsBeforeWarning", options.MaxDependentsBeforeWarning);
        options.TraceMaxDepth = ReadInt(json, "traceMaxDepth", options.TraceMaxDepth);
        options.TraceSafetyLimit = ReadInt(json, "traceSafetyLimit", options.TraceSafetyLimit);
        options.ConfirmLargeDependentScan = ReadBool(json, "confirmLargeDependentScan", options.ConfirmLargeDependentScan);
        options.ExplorerWindow = ReadPlacement(json, "explorerWindow") ?? options.ExplorerWindow;
        options.DependentsWindow = ReadPlacement(json, "dependentsWindow") ?? options.DependentsWindow;
        return options;
    }

    private static void WritePlacement(StringBuilder builder, string name, WindowPlacement placement, bool comma)
    {
        builder.Append("  \"").Append(name).Append("\": {");
        builder.Append(" \"left\": ").Append(FormatNumber(placement.Left)).Append(',');
        builder.Append(" \"top\": ").Append(FormatNumber(placement.Top)).Append(',');
        builder.Append(" \"width\": ").Append(FormatNumber(placement.Width)).Append(',');
        builder.Append(" \"height\": ").Append(FormatNumber(placement.Height));
        builder.Append(" }");
        if (comma) builder.Append(',');
        builder.AppendLine();
    }

    private static WindowPlacement? ReadPlacement(string json, string name)
    {
        var match = Regex.Match(json, "\"" + Regex.Escape(name) + "\"\\s*:\\s*\\{([^}]*)\\}");
        if (!match.Success) return null;
        var body = match.Groups[1].Value;
        return new WindowPlacement
        {
            Left = ReadNumber(body, "left"),
            Top = ReadNumber(body, "top"),
            Width = ReadNumber(body, "width"),
            Height = ReadNumber(body, "height")
        };
    }

    private static void WriteString(StringBuilder builder, string name, string value, bool comma)
    {
        builder.Append("  \"").Append(name).Append("\": \"").Append(Escape(value)).Append('"');
        if (comma) builder.Append(',');
        builder.AppendLine();
    }

    private static void WriteInt(StringBuilder builder, string name, int value, bool comma)
    {
        builder.Append("  \"").Append(name).Append("\": ").Append(value);
        if (comma) builder.Append(',');
        builder.AppendLine();
    }

    private static void WriteBool(StringBuilder builder, string name, bool value, bool comma)
    {
        builder.Append("  \"").Append(name).Append("\": ").Append(value ? "true" : "false");
        if (comma) builder.Append(',');
        builder.AppendLine();
    }

    private static string? ReadString(string json, string name)
    {
        var match = Regex.Match(json, "\"" + Regex.Escape(name) + "\"\\s*:\\s*\"((?:\\\\.|[^\"\\\\])*)\"");
        return match.Success ? Unescape(match.Groups[1].Value) : null;
    }

    private static int ReadInt(string json, string name, int fallback)
    {
        var match = Regex.Match(json, "\"" + Regex.Escape(name) + "\"\\s*:\\s*(-?\\d+)");
        return match.Success && int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : fallback;
    }

    private static bool ReadBool(string json, string name, bool fallback)
    {
        var match = Regex.Match(json, "\"" + Regex.Escape(name) + "\"\\s*:\\s*(true|false)", RegexOptions.IgnoreCase);
        return match.Success ? string.Equals(match.Groups[1].Value, "true", StringComparison.OrdinalIgnoreCase) : fallback;
    }

    private static double ReadNumber(string json, string name)
    {
        var match = Regex.Match(json, "\"" + Regex.Escape(name) + "\"\\s*:\\s*(-?\\d+(?:\\.\\d+)?)");
        return match.Success && double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0;
    }

    private static string FormatNumber(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);

    private static string Escape(string value) =>
        (value ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"");

    private static string Unescape(string value) =>
        value.Replace("\\\"", "\"").Replace("\\\\", "\\");
}
