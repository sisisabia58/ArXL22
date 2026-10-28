using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Controls;

namespace ArixcelExplorer.UI.Services;

/// <summary>
/// Remembers GridView column widths for the process and in
/// %AppData%\EXLerate\columns.txt so the next Excel session starts from the drag.
/// </summary>
internal static class ExplorerColumnLayout
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "EXLerate",
        "columns.txt");

    private static readonly Dictionary<string, double[]> Cache =
        new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);

    private static bool _loaded;

    public static void Attach(GridView view, string key, params double[] defaults)
    {
        if (view == null || defaults == null || defaults.Length == 0) return;
        EnsureLoaded();
        var tracked = defaults.Length;
        var widths = Cache.TryGetValue(key, out var saved) && saved.Length == tracked
            ? saved
            : defaults;
        for (var i = 0; i < tracked && i < view.Columns.Count; i++)
        {
            var width = widths[i];
            if (width < 36 || double.IsNaN(width) || double.IsInfinity(width)) width = defaults[i];
            view.Columns[i].Width = width;
        }

        var descriptor = DependencyPropertyDescriptor.FromProperty(
            GridViewColumn.WidthProperty,
            typeof(GridViewColumn));
        for (var i = 0; i < tracked && i < view.Columns.Count; i++)
        {
            descriptor.AddValueChanged(view.Columns[i], (_, _) => Remember(key, view, tracked));
        }
    }

    private static void Remember(string key, GridView view, int tracked)
    {
        var widths = new double[tracked];
        for (var i = 0; i < tracked && i < view.Columns.Count; i++)
        {
            // Width is the dragged size. ActualWidth on the old last column included
            // the viewport stretch and would pin Count to the right on the next open.
            var width = view.Columns[i].Width;
            widths[i] = width >= 36 && !double.IsNaN(width) && !double.IsInfinity(width) ? width : 36;
        }

        Cache[key] = widths;
        try
        {
            var directory = Path.GetDirectoryName(FilePath);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
            var builder = new StringBuilder();
            foreach (var pair in Cache)
            {
                builder.Append(pair.Key);
                builder.Append('=');
                for (var i = 0; i < pair.Value.Length; i++)
                {
                    if (i > 0) builder.Append(',');
                    builder.Append(pair.Value[i].ToString("0.##", CultureInfo.InvariantCulture));
                }

                builder.AppendLine();
            }

            File.WriteAllText(FilePath, builder.ToString());
        }
        catch
        {
            // Column widths still stay in memory for this Excel session.
        }
    }

    private static void EnsureLoaded()
    {
        if (_loaded) return;
        _loaded = true;
        try
        {
            if (!File.Exists(FilePath)) return;
            foreach (var line in File.ReadAllLines(FilePath))
            {
                var split = line.Split(new[] { '=' }, 2);
                if (split.Length != 2) continue;
                var parts = split[1].Split(',');
                var widths = new double[parts.Length];
                var ok = true;
                for (var i = 0; i < parts.Length; i++)
                {
                    if (!double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out widths[i]))
                    {
                        ok = false;
                        break;
                    }
                }

                if (ok && widths.Length > 0) Cache[split[0].Trim()] = widths;
            }
        }
        catch
        {
            // Defaults are used when the file cannot be read.
        }
    }
}
