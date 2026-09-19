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
    private readonly Stack<ExplorerStackEntry> _stack = new();

    public void Push(ExplorerStackEntry entry) => _stack.Push(entry);

    public ExplorerStackEntry? Pop() => _stack.Count > 0 ? _stack.Pop() : null;

    public ExplorerStackEntry? Peek() => _stack.Count > 0 ? _stack.Peek() : null;

    public void Clear() => _stack.Clear();
}
