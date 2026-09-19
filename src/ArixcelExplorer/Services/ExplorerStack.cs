using System.Collections.Generic;

namespace ArixcelExplorer.Services;

public sealed class ExplorerStackEntry
{
    public string OriginAddress { get; set; } = "";
    public string WorksheetName { get; set; } = "";
    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
}

public sealed class ExplorerStack
{
    private readonly List<ExplorerStackEntry> _items = new();

    public int Count => _items.Count;

    public ExplorerStackEntry? First => _items.Count > 0 ? _items[0] : null;

    public void Push(ExplorerStackEntry entry) => _items.Add(entry);

    public ExplorerStackEntry? Pop()
    {
        if (_items.Count == 0) return null;
        var last = _items[_items.Count - 1];
        _items.RemoveAt(_items.Count - 1);
        return last;
    }

    public ExplorerStackEntry? Peek() => _items.Count > 0 ? _items[_items.Count - 1] : null;

    public void Clear() => _items.Clear();
}
