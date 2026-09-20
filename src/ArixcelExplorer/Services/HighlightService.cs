using System;
using System.Collections.Generic;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.Services;

public sealed class HighlightService
{
    private readonly Excel.Application _app;
    private readonly Dictionary<string, int> _original = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<OverlayLayer>> _layers = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, HashSet<string>> _ownerCells = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<string>> _transient = new(StringComparer.OrdinalIgnoreCase);

    public HighlightService(Excel.Application application)
    {
        _app = application;
    }

    public void Apply(string ownerId, string address, string hex)
    {
        var range = Resolve(address);
        if (range == null) return;
        var color = ExcelOleColor.FromHex(hex);
        EnsureOwner(ownerId);

        foreach (Excel.Range cell in range.Cells)
        {
            var key = KeyOf(cell);
            if (!_original.ContainsKey(key))
            {
                _original[key] = Convert.ToInt32(cell.Interior.Color);
            }

            _ownerCells[ownerId].Add(key);
            var layers = LayersFor(key);
            layers.RemoveAll(layer => string.Equals(layer.OwnerId, ownerId, StringComparison.OrdinalIgnoreCase));
            layers.Add(new OverlayLayer { OwnerId = ownerId, Color = color });
            cell.Interior.Color = color;
        }
    }

    public void ApplyMany(string ownerId, IEnumerable<string> addresses, string hex)
    {
        foreach (var address in addresses)
        {
            if (!string.IsNullOrWhiteSpace(address))
            {
                Apply(ownerId, address, hex);
            }
        }
    }

    public void SetTransient(string ownerId, string? address, string hex, string originAddress)
    {
        SetTransientMany(ownerId, address == null ? Array.Empty<string>() : new[] { address }, hex, originAddress);
    }

    public void SetTransientMany(string ownerId, IEnumerable<string> addresses, string hex, string originAddress)
    {
        if (_transient.TryGetValue(ownerId, out var previousList))
        {
            foreach (var previous in previousList)
            {
                if (!AddressesEqual(previous, originAddress))
                {
                    ReleaseAddress(ownerId, previous);
                }
            }

            _transient.Remove(ownerId);
        }

        var applied = new List<string>();
        foreach (var address in addresses)
        {
            if (string.IsNullOrWhiteSpace(address) || AddressesEqual(address, originAddress))
            {
                continue;
            }

            Apply(ownerId, address, hex);
            applied.Add(address);
        }

        if (applied.Count > 0)
        {
            _transient[ownerId] = applied;
        }
    }

    public void ReleaseOwner(string ownerId)
    {
        if (!_ownerCells.TryGetValue(ownerId, out var keys)) return;
        foreach (var key in keys.ToArraySafe())
        {
            RemoveLayer(ownerId, key);
        }

        _ownerCells.Remove(ownerId);
        _transient.Remove(ownerId);
    }

    public void RestoreAll()
    {
        foreach (var pair in _original)
        {
            try
            {
                var range = Resolve(pair.Key);
                if (range != null) range.Interior.Color = pair.Value;
            }
            catch
            {
                // ignore restore failures
            }
        }

        _original.Clear();
        _layers.Clear();
        _ownerCells.Clear();
        _transient.Clear();
    }

    private void ReleaseAddress(string ownerId, string address)
    {
        var range = Resolve(address);
        if (range == null) return;
        foreach (Excel.Range cell in range.Cells)
        {
            RemoveLayer(ownerId, KeyOf(cell));
            if (_ownerCells.TryGetValue(ownerId, out var keys))
            {
                keys.Remove(KeyOf(cell));
            }
        }
    }

    private void RemoveLayer(string ownerId, string key)
    {
        if (_layers.TryGetValue(key, out var layers))
        {
            layers.RemoveAll(layer => string.Equals(layer.OwnerId, ownerId, StringComparison.OrdinalIgnoreCase));
        }

        var range = Resolve(key);
        if (range == null) return;

        if (!_layers.TryGetValue(key, out var remaining) || remaining.Count == 0)
        {
            if (_original.TryGetValue(key, out var original))
            {
                range.Interior.Color = original;
                _original.Remove(key);
            }

            _layers.Remove(key);
            return;
        }

        range.Interior.Color = remaining[remaining.Count - 1].Color;
    }

    private void EnsureOwner(string ownerId)
    {
        if (!_ownerCells.ContainsKey(ownerId))
        {
            _ownerCells[ownerId] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private List<OverlayLayer> LayersFor(string key)
    {
        if (!_layers.TryGetValue(key, out var layers))
        {
            layers = new List<OverlayLayer>();
            _layers[key] = layers;
        }

        return layers;
    }

    private Excel.Range? Resolve(string address)
    {
        if (string.IsNullOrWhiteSpace(address)) return null;
        try
        {
            var parsed = TraceUtils.ParseWorksheetScopedAddress(address);
            if (parsed != null)
            {
                var sheet = FindWorksheet(parsed.WorksheetName);
                return sheet?.Range[parsed.RangeAddress];
            }

            return _app.Range[address];
        }
        catch
        {
            return null;
        }
    }

    private Excel.Worksheet? FindWorksheet(string name)
    {
        foreach (Excel.Worksheet ws in _app.Worksheets)
        {
            if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return ws;
            }
        }

        return null;
    }

    private static string KeyOf(Excel.Range cell)
    {
        var ws = cell.Worksheet as Excel.Worksheet;
        return $"'{ws?.Name}'!{cell.Address[false, false]}";
    }

    private static bool AddressesEqual(string? left, string? right) =>
        string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

    private sealed class OverlayLayer
    {
        public string OwnerId { get; set; } = "";
        public int Color { get; set; }
    }
}

internal static class HighlightCollectionExtensions
{
    public static string[] ToArraySafe(this HashSet<string> source)
    {
        var copy = new string[source.Count];
        source.CopyTo(copy);
        return copy;
    }
}
