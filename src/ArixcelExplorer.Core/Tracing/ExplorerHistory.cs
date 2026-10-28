using System;
using System.Collections.Generic;

namespace ArixcelExplorer.Core.Tracing;

/// <summary>
/// Exploration addresses for Go Back. Arixcel keeps up to 100 and wraps after the oldest.
/// </summary>
public sealed class ExplorerHistory
{
    public const int Capacity = 100;

    private readonly List<string> _addresses = new();
    private int _cursor = -1;

    public int Count => _addresses.Count;

    public void Record(string? address)
    {
        if (string.IsNullOrWhiteSpace(address)) return;
        var value = address!.Trim();
        address = value;
        if (_cursor >= 0 && _cursor < _addresses.Count - 1)
        {
            _addresses.RemoveRange(_cursor + 1, _addresses.Count - _cursor - 1);
        }

        if (_addresses.Count > 0 &&
            _cursor == _addresses.Count - 1 &&
            string.Equals(_addresses[_addresses.Count - 1], address, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _addresses.Add(address);
        while (_addresses.Count > Capacity)
        {
            _addresses.RemoveAt(0);
        }

        _cursor = _addresses.Count - 1;
    }

    public string? Back()
    {
        if (_addresses.Count == 0) return null;
        if (_cursor <= 0)
        {
            _cursor = _addresses.Count - 1;
            return _addresses[_cursor];
        }

        _cursor -= 1;
        return _addresses[_cursor];
    }
}
