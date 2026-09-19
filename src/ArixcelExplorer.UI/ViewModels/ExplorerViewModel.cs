using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.UI.ViewModels;

public enum ExplorerCloseMode
{
    KeepSelection,
    RestorePrevious
}

public sealed class ExplorerTreeRow : INotifyPropertyChanged
{
    public string Id { get; set; } = "";
    public int Indent { get; set; }
    public string Component { get; set; } = "";
    public string Info { get; set; } = "";
    public string Value { get; set; } = "";
    public string Location { get; set; } = "";
    public bool IsActiveBranch { get; set; }
    public bool HasChildren { get; set; }
    public bool IsExpanded { get; set; }

    private bool _isSelected;
    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;
            _isSelected = value;
            OnPropertyChanged();
        }
    }

    public string ComponentPrefix
    {
        get
        {
            var indent = new string(' ', Indent * 2);
            var glyph = HasChildren ? (IsExpanded ? "-" : "+") : " ";
            return indent + glyph + " ";
        }
    }

    public string ComponentDisplay => ComponentPrefix + Component;

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public sealed class ExplorerViewModel : INotifyPropertyChanged
{
    public ObservableCollection<ExplorerTreeRow> Rows { get; } = new();

    private string _formulaText = "";
    public string FormulaText
    {
        get => _formulaText;
        set
        {
            if (_formulaText == value) return;
            _formulaText = value;
            OnPropertyChanged();
        }
    }

    private string _rootAddress = "";
    public string RootAddress
    {
        get => _rootAddress;
        set
        {
            if (_rootAddress == value) return;
            _rootAddress = value;
            OnPropertyChanged();
        }
    }

    private string _statusText = "";
    public string StatusText
    {
        get => _statusText;
        set
        {
            if (_statusText == value) return;
            _statusText = value;
            OnPropertyChanged();
        }
    }

    private FormulaAstNode? _root;
    private int _selectedIndex = -1;
    private bool _isFullyExpanded = true;

    public int SelectedIndex
    {
        get => _selectedIndex;
        set
        {
            if (value < 0) return;
            SelectRow(value);
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event System.Action<ExplorerTreeRow>? NavigateRequested;
    public event System.Action<ExplorerCloseMode>? CloseRequested;

    public void LoadTree(FormulaAstNode root, string formulaText)
    {
        _root = root;
        _isFullyExpanded = true;
        FormulaText = formulaText;
        RefreshRows();
        if (Rows.Count > 0) SelectRow(0);
    }

    public void RefreshRows()
    {
        var selectedId = _selectedIndex >= 0 && _selectedIndex < Rows.Count ? Rows[_selectedIndex].Id : null;
        Rows.Clear();
        if (_root == null) return;

        var depthById = new System.Collections.Generic.Dictionary<string, int>();
        void Walk(FormulaAstNode node, int depth)
        {
            depthById[node.Id] = depth;
            if (!node.IsExpanded) return;
            foreach (var child in node.Children) Walk(child, depth + 1);
        }

        Walk(_root, 0);

        foreach (var node in FormulaAstParser.FlattenVisible(_root))
        {
                var rawLocation = node.Location ?? "";
                var location = string.IsNullOrWhiteSpace(rawLocation)
                    ? ""
                    : TraceUtils.QualifyAddress(rawLocation, RootAddress);
            var component = node.Label;
            // Origin row: drop 'Sheet'! prefix from Element when label already equals Location.
            if (node.Kind == FormulaNodeKind.Root &&
                string.Equals(node.Label, node.Location, System.StringComparison.OrdinalIgnoreCase))
            {
                var parsed = TraceUtils.ParseWorksheetScopedAddress(node.Label);
                if (parsed != null)
                {
                    component = parsed.RangeAddress;
                }
            }

            Rows.Add(new ExplorerTreeRow
            {
                Id = node.Id,
                Indent = depthById.TryGetValue(node.Id, out var depth) ? depth : 0,
                Component = component,
                Info = node.Info,
                Value = node.Value ?? "",
                Location = location,
                IsActiveBranch = node.IsActiveBranch,
                HasChildren = node.Children.Count > 0,
                IsExpanded = node.IsExpanded,
                IsSelected = false
            });
        }

        if (selectedId != null)
        {
            var restored = IndexOfId(selectedId);
            if (restored >= 0)
            {
                ApplySelection(restored, navigate: false);
            }
        }
    }

    public void SelectRow(int index)
    {
        if (index < 0 || index >= Rows.Count) return;
        ApplySelection(index, navigate: true);
    }

    public void MoveSelection(int delta)
    {
        if (Rows.Count == 0) return;
        var next = _selectedIndex < 0 ? 0 : System.Math.Max(0, System.Math.Min(_selectedIndex + delta, Rows.Count - 1));
        SelectRow(next);
    }

    public void ExpandSelected()
    {
        if (_selectedIndex < 0 || _root == null) return;
        var row = Rows[_selectedIndex];
        var node = FindNode(_root, row.Id);
        if (node == null || node.Children.Count == 0 || node.IsExpanded) return;
        node.IsExpanded = true;
        RefreshRows();
        SelectById(row.Id);
    }

    public void CollapseSelectedOrMoveToParent()
    {
        if (_selectedIndex < 0 || _root == null) return;
        var row = Rows[_selectedIndex];
        var node = FindNode(_root, row.Id);
        if (node != null && node.Children.Count > 0 && node.IsExpanded)
        {
            node.IsExpanded = false;
            RefreshRows();
            SelectById(row.Id);
            return;
        }

        for (var i = _selectedIndex - 1; i >= 0; i--)
        {
            if (Rows[i].Indent < row.Indent)
            {
                SelectRow(i);
                return;
            }
        }
    }

    public void CycleExpandCollapse()
    {
        if (_root == null) return;
        _isFullyExpanded = FormulaAstParser.CycleExpandAll(_root, _isFullyExpanded);
        var selectedId = _selectedIndex >= 0 && _selectedIndex < Rows.Count ? Rows[_selectedIndex].Id : null;
        RefreshRows();
        if (selectedId != null) SelectById(selectedId);
        else if (Rows.Count > 0) SelectRow(0);
        StatusText = _isFullyExpanded ? "Expanded all" : "Collapsed all";
    }

    public void ExpandAll()
    {
        if (_root == null) return;
        FormulaAstParser.ExpandAll(_root);
        _isFullyExpanded = true;
        RefreshRows();
    }

    public void CollapseAll()
    {
        if (_root == null) return;
        FormulaAstParser.CollapseAll(_root);
        _isFullyExpanded = false;
        RefreshRows();
    }

    public void RequestKeepClose() => CloseRequested?.Invoke(ExplorerCloseMode.KeepSelection);

    public void RequestBackClose() => CloseRequested?.Invoke(ExplorerCloseMode.RestorePrevious);

    private void SelectById(string id)
    {
        var index = IndexOfId(id);
        if (index >= 0) SelectRow(index);
        else if (Rows.Count > 0) SelectRow(System.Math.Min(_selectedIndex, Rows.Count - 1));
    }

    private int IndexOfId(string id)
    {
        for (var i = 0; i < Rows.Count; i++)
        {
            if (Rows[i].Id == id) return i;
        }

        return -1;
    }

    private void ApplySelection(int index, bool navigate)
    {
        for (var i = 0; i < Rows.Count; i++) Rows[i].IsSelected = i == index;
        if (_selectedIndex != index)
        {
            _selectedIndex = index;
            OnPropertyChanged(nameof(SelectedIndex));
        }
        else
        {
            _selectedIndex = index;
        }

        if (navigate)
        {
            NavigateRequested?.Invoke(Rows[index]);
        }
    }

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
