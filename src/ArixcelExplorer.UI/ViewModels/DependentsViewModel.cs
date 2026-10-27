using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.UI.ViewModels;

public sealed class DependentRow : INotifyPropertyChanged
{
    public string Key { get; set; } = "";
    public string Address { get; set; } = "";
    public string Display { get; set; } = "";
    public int Count { get; set; }
    public string Value { get; set; } = "";
    public bool IsOrigin { get; set; }
    public bool IsGroup { get; set; }
    public bool HasChildren { get; set; }
    public bool IsExpanded { get; set; }
    public int Indent { get; set; }
    public IList<ExplorerTreeGuide> Guides { get; set; } = new List<ExplorerTreeGuide>();

    public string CountDisplay => HasChildren || IsOrigin ? Count.ToString() : "";
    public bool ShowDropLine => HasChildren && IsExpanded;

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

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public sealed class DependentsViewModel : INotifyPropertyChanged
{
    private readonly List<DependentEntry> _entries = new();
    private DependentsTreeNode? _root;
    private string _originAddress = "";
    private string _originValue = "";
    private bool _groupBySheet;
    private bool _rebuilding;
    private int _selectedIndex = -1;

    public ObservableCollection<DependentRow> Rows { get; } = new();

    public string OriginAddress => _originAddress;

    public IReadOnlyList<DependentEntry> Entries => _entries;

    private string _sourceSummary = "";
    public string SourceSummary
    {
        get => _sourceSummary;
        set
        {
            if (_sourceSummary == value) return;
            _sourceSummary = value;
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

    public bool GroupBySheet
    {
        get => _groupBySheet;
        set
        {
            if (_groupBySheet == value) return;
            _groupBySheet = value;
            OnPropertyChanged();
            Rebuild();
        }
    }

    public int SelectedIndex
    {
        get => _selectedIndex;
        set
        {
            if (value < 0 || _rebuilding) return;
            SelectRow(value);
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action<DependentRow>? NavigateRequested;
    public event Action<ExplorerCloseMode>? CloseRequested;
    public event Action? DrillRequested;
    public event Action? RefreshRequested;
    public event Action? GoToExcelRequested;

    public void Load(IReadOnlyList<DependentEntry> entries, string originAddress, string originValue = "")
    {
        _originAddress = originAddress ?? "";
        _originValue = originValue ?? "";
        SourceSummary = _originAddress;
        ReplaceEntries(entries);
        Rebuild();
    }

    public void ApplyAreas(IReadOnlyList<DependentEntry> entries)
    {
        ReplaceEntries(entries);
        Rebuild();
    }

    public void SelectRow(int index, bool navigate = true)
    {
        if (index < 0 || index >= Rows.Count) return;
        for (var i = 0; i < Rows.Count; i++) Rows[i].IsSelected = i == index;
        if (_selectedIndex != index)
        {
            _selectedIndex = index;
            OnPropertyChanged(nameof(SelectedIndex));
        }

        if (!navigate || _rebuilding) return;
        var row = Rows[index];
        if (!row.IsGroup && !string.IsNullOrWhiteSpace(row.Address))
        {
            NavigateRequested?.Invoke(row);
        }
    }

    public void MoveSelection(int delta)
    {
        if (Rows.Count == 0) return;
        var next = _selectedIndex < 0 ? 0 : Math.Max(0, Math.Min(_selectedIndex + delta, Rows.Count - 1));
        SelectRow(next);
    }

    public void ExpandOrMove(int delta)
    {
        var row = SelectedRow;
        if (row == null)
        {
            MoveSelection(delta);
            return;
        }

        if (delta > 0 && row.HasChildren && !row.IsExpanded)
        {
            ToggleExpand(row.Key);
            return;
        }

        if (delta < 0 && row.HasChildren && row.IsExpanded)
        {
            ToggleExpand(row.Key);
            return;
        }

        MoveSelection(delta);
    }

    public void ToggleExpand(string? key)
    {
        if (_root == null || string.IsNullOrEmpty(key)) return;
        var node = Find(_root, key ?? "");
        if (node == null || node.Children.Count == 0) return;
        node.IsExpanded = !node.IsExpanded;
        Flatten(key);
    }

    public void RequestKeepClose() => CloseRequested?.Invoke(ExplorerCloseMode.KeepSelection);

    public void RequestBackClose() => CloseRequested?.Invoke(ExplorerCloseMode.RestorePrevious);

    public void RequestRefresh() => RefreshRequested?.Invoke();

    public void RequestGoToExcel() => GoToExcelRequested?.Invoke();

    public void SelectOrigin() => SelectRow(0);

    public void RequestDrill()
    {
        var row = SelectedRow;
        if (row == null || row.IsOrigin || row.IsGroup || string.IsNullOrWhiteSpace(row.Address)) return;
        DrillRequested?.Invoke();
    }

    public void HandleCtrlShiftQ() => RequestDrill();

    public DependentRow? SelectedRow =>
        _selectedIndex >= 0 && _selectedIndex < Rows.Count ? Rows[_selectedIndex] : null;

    private void ReplaceEntries(IReadOnlyList<DependentEntry>? entries)
    {
        _entries.Clear();
        if (entries == null) return;
        foreach (var entry in entries)
        {
            if (entry == null) continue;
            _entries.Add(entry);
        }
    }

    private void Rebuild()
    {
        _root = DependentsTree.Build(_originAddress, _originValue, _entries, _groupBySheet);
        Flatten(null);
    }

    private void Flatten(string? reselectKey)
    {
        _rebuilding = true;
        try
        {
            Rows.Clear();
            if (_root != null)
            {
                Append(_root, 0, Array.Empty<bool>(), true);
            }

            var index = 0;
            if (!string.IsNullOrEmpty(reselectKey))
            {
                for (var i = 0; i < Rows.Count; i++)
                {
                    if (Rows[i].Key == reselectKey)
                    {
                        index = i;
                        break;
                    }
                }
            }

            _selectedIndex = -1;
            if (Rows.Count > 0)
            {
                SelectRow(Math.Min(index, Rows.Count - 1), navigate: false);
            }

            var dependents = _root?.Count ?? 0;
            StatusText = dependents <= 0 ? ExplorerLocations.DependentsStatus(0) : "";
        }
        finally
        {
            _rebuilding = false;
        }
    }

    private void Append(DependentsTreeNode node, int depth, bool[] ancestorContinues, bool isLast)
    {
        var key = node.IsRoot
            ? "root"
            : node.IsGroup
                ? "group:" + node.SheetName
                : "cell:" + node.Address;
        var guides = new List<ExplorerTreeGuide>(depth);
        for (var level = 0; level < depth; level++)
        {
            var elbow = level == depth - 1;
            guides.Add(new ExplorerTreeGuide
            {
                FullVertical = !elbow && ancestorContinues[level],
                Elbow = elbow,
                ElbowContinues = elbow && !isLast
            });
        }

        Rows.Add(new DependentRow
        {
            Key = key,
            Address = node.Address,
            Display = node.Display,
            Count = node.Count,
            Value = node.Value ?? "",
            IsOrigin = node.IsRoot,
            IsGroup = node.IsGroup,
            HasChildren = node.Children.Count > 0,
            IsExpanded = node.IsExpanded,
            Indent = depth,
            Guides = guides
        });

        if (!node.IsExpanded || node.Children.Count == 0) return;

        var childContinues = new bool[depth + 1];
        Array.Copy(ancestorContinues, childContinues, depth);
        childContinues[depth] = !isLast;
        for (var i = 0; i < node.Children.Count; i++)
        {
            Append(node.Children[i], depth + 1, childContinues, i == node.Children.Count - 1);
        }
    }

    private static DependentsTreeNode? Find(DependentsTreeNode node, string key)
    {
        var nodeKey = node.IsRoot
            ? "root"
            : node.IsGroup
                ? "group:" + node.SheetName
                : "cell:" + node.Address;
        if (nodeKey == key) return node;
        foreach (var child in node.Children)
        {
            var found = Find(child, key);
            if (found != null) return found;
        }

        return null;
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
