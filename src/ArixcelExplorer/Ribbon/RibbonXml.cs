using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using ArixcelExplorer.Services;

namespace ArixcelExplorer.Ribbon;

internal static class RibbonXml
{
    private const string MutexName = @"Local\EXLerateExplorer.RibbonXml";
    private const string Empty = """
        <?xml version="1.0" encoding="UTF-8"?>
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" />
        """;

    private static Mutex? _claim;

    public static string Load(string host, string ribbonId)
    {
        try
        {
            if (!TryClaim())
            {
                AddInLog.Info($"GetCustomUI skipped host={host} ribbonId={ribbonId}");
                return Empty;
            }

            var xml = Read();
            AddInLog.Info($"GetCustomUI served host={host} ribbonId={ribbonId} length={xml.Length}");
            return xml;
        }
        catch (Exception ex)
        {
            AddInLog.Error(ex);
            return Empty;
        }
    }

    private static bool TryClaim()
    {
        if (_claim != null) return true;

        try
        {
            var mutex = new Mutex(true, MutexName, out var created);
            if (created)
            {
                _claim = mutex;
                return true;
            }

            mutex.Dispose();
            return false;
        }
        catch (AbandonedMutexException ex)
        {
            _claim = ex.Mutex;
            return true;
        }
        catch
        {
            return true;
        }
    }

    private static string Read()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = assembly.GetManifestResourceNames()
            .FirstOrDefault(name => name.EndsWith("ArixcelRibbon.xml", StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrEmpty(resourceName))
        {
            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream != null)
            {
                using var reader = new StreamReader(stream);
                var xml = reader.ReadToEnd();
                if (!string.IsNullOrWhiteSpace(xml)) return xml;
            }
        }

        var path = Path.Combine(Path.GetDirectoryName(assembly.Location) ?? "", "Ribbon", "ArixcelRibbon.xml");
        if (File.Exists(path)) return File.ReadAllText(path);

        AddInLog.Info("GetCustomUI fell back to empty ribbon XML");
        return Empty;
    }
}
