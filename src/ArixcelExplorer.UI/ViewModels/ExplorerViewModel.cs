using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ArixcelExplorer.Core.Formulas;

namespace ArixcelExplorer.UI.ViewModels;

public sealed class ExplorerTreeRow : INotifyPropertyChanged
{
    public string Id { get; set; } = "";
    public int Indent { get; set; }
    public string Component { get; set; } = "";
    public string Value { get; set; } = "";
    public string Location { get; set; } = "";
    public bool IsActiveBranch { get; set; }
    public bool HasChildren { get; set; }
    public bool IsExpanded { get; set; }
    public bool IsSelected { get; set; }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public sealed class ExplorerViewModel : INotifyPropertyChanged
{
    public ObservableCollection<ExplorerTreeRow> Rows { get; } = new();
    public string FormulaText { get; set; } = "";
    public string RootAddress { get; set; } = "";
    public string StatusText { get; set; } = "";

    private FormulaAstNode? _root;
    private int _selectedIndex = -1;

    public event PropertyChangedEventHandler? PropertyChanged;
    public event System.Action<ExplorerTreeRow>? NavigateRequested;
    public event System.Action? CloseRequested;
    public event System.Action? DrillDownRequested;

    public void LoadTree(FormulaAstNode root, string formulaText)
    {
        _root = root;
        FormulaText = formulaText;
        RefreshRows();
    }

    public void RefreshRows()
    {
        Rows.Clear();
        if (_root == null) return;

        var flat = FormulaAstParser.FlattenVisible(_root);
        var depthById = new System.Collections.Generic.Dictionary<string, int>();
        void Walk(FormulaAstNode node, int depth)
        {
            depthById[node.Id] = depth;
            if (!node.IsExpanded) return;
            foreach (var child in node.Children) Walk(child, depth + 1);
        }

        Walk(_root, 0);

        foreach (var node in flat)
        {
            if (node.Kind == FormulaNodeKind.Root) continue;
            Rows.Add(new ExplorerTreeRow
            {
                Id = node.Id,
                Indent = depthById.TryGetValue(node.Id, out var depth) ? depth : 0,
                Component = node.Label,
                Value = node.Value ?? "",
                Location = node.Location ?? "",
                IsActiveBranch = node.IsActiveBranch,
                HasChildren = node.Children.Count > 0,
                IsExpanded = node.IsExpanded,
                IsSelected = false
            });
        }
    }

    public void SelectRow(int index)
    {
        if (index < 0 || index >= Rows.Count) return;
        for (var i = 0; i < Rows.Count; i++) Rows[i].IsSelected = i == index;
        _selectedIndex = index;
        NavigateRequested?.Invoke(Rows[index]);
    }

    public void MoveSelection(int delta)
    {
        if (Rows.Count == 0) return;
        var next = _selectedIndex < 0 ? 0 : System.Math.Max(0, System.Math.Min(_selectedIndex + delta, Rows.Count - 1));
        SelectRow(next);
    }

    public void ToggleExpandSelected()
    {
        if (_selectedIndex < 0 || _root == null) return;
        var row = Rows[_selectedIndex];
        var node = FindNode(_root, row.Id);
        if (node == null || node.Children.Count == 0) return;
        FormulaAstParser.ToggleExpand(node);
        RefreshRows();
        SelectRow(System.Math.Min(_selectedIndex, Rows.Count - 1));
    }

    public void ExpandAll()
    {
        if (_root == null) return;
        FormulaAstParser.ExpandAll(_root);
        RefreshRows();
    }

    public void CollapseAll()
    {
        if (_root == null) return;
        FormulaAstParser.CollapseAll(_root);
        RefreshRows();
    }

    public void RequestDrillDown() => DrillDownRequested?.Invoke();

    public void RequestClose() => CloseRequested?.Invoke();

    private static FormulaAstNode? FindNode(FormulaAstNode root, string id)
    {
        if (root.Id == id) return root;
        foreach (var child in root.Children)
        {
            var found = FindNode(child, id);
            if (found != null) return found;
        }

        return null;
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
