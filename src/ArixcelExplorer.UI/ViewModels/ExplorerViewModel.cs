using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.UI.ViewModels;

public enum ExplorerCloseMode
{
    KeepSelection,
    RestorePrevious
}

public sealed class ExplorerTreeGuide
{
    public bool FullVertical { get; set; }
    public bool Elbow { get; set; }
    public bool ElbowContinues { get; set; }
}

public sealed class ExplorerTreeRow : INotifyPropertyChanged
{
    public string Id { get; set; } = "";
    public int Indent { get; set; }
    public string Component { get; set; } = "";
    public string Info { get; set; } = "";
    public string Value { get; set; } = "";
    public string Location { get; set; } = "";
    public string NavigationAddress { get; set; } = "";
    public bool IsActiveBranch { get; set; }
    public bool HasChildren { get; set; }
    public bool IsExpanded { get; set; }
    public int SourceStart { get; set; }
    public int SourceLength { get; set; }
    public FormulaNodeKind Kind { get; set; }

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

    public string Glyph => ExplorerTreeChrome.Glyph(HasChildren, IsExpanded);

    public bool ShowDropLine => HasChildren && IsExpanded;

    public double IndentWidth => ExplorerTreeChrome.IndentWidth(Indent);

    public System.Collections.Generic.IList<ExplorerTreeGuide> Guides { get; set; } =
        new System.Collections.Generic.List<ExplorerTreeGuide>();

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

    private bool _isEditingFormula;
    public bool IsEditingFormula
    {
        get => _isEditingFormula;
        private set
        {
            if (_isEditingFormula == value) return;
            _isEditingFormula = value;
            OnPropertyChanged();
        }
    }

    public string? FormulaEditSeed { get; private set; }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event System.Action<ExplorerTreeRow>? NavigateRequested;
    public event System.Action<ExplorerCloseMode>? CloseRequested;
    public event System.Action? DrillRequested;
    public event System.Action? RefreshRequested;
    public event System.Action? ExplorePrecedentsRequested;
    public event System.Action? ExploreDependentsRequested;
    public event System.Action? GoToExcelRequested;
    public event System.Action<string>? FormulaApplyRequested;

    public void LoadTree(FormulaAstNode root, string formulaText)
    {
        _root = root;
        _isFullyExpanded = true;
        FormulaEditSeed = null;
        IsEditingFormula = false;
        FormulaAstParser.CollapseFunctions(root);
        FormulaText = formulaText;
        RefreshRows();
        if (Rows.Count > 0)
        {
            ApplySelection(ExplorerLocations.OpeningSelectionIndex(Rows.Count), navigate: false);
        }
    }

    public void RefreshRows()
    {
        var selectedId = _selectedIndex >= 0 && _selectedIndex < Rows.Count ? Rows[_selectedIndex].Id : null;
        Rows.Clear();
        if (_root == null) return;

        AppendNode(_root, 0, System.Array.Empty<bool>(), isLast: true);

        if (selectedId != null)
        {
            var restored = IndexOfId(selectedId);
            if (restored >= 0)
            {
                ApplySelection(restored, navigate: false);
            }
        }
    }

    public void SelectRow(int index, bool navigate = true)
    {
        if (index < 0 || index >= Rows.Count) return;
        ApplySelection(index, navigate);
    }

    public void MoveSelection(int delta)
    {
        if (Rows.Count == 0) return;
        var next = _selectedIndex < 0 ? 0 : System.Math.Max(0, System.Math.Min(_selectedIndex + delta, Rows.Count - 1));
        SelectRow(next);
    }

    public void ToggleExpand(string id)
    {
        if (_root == null || string.IsNullOrWhiteSpace(id)) return;
        var node = FindNode(_root, id);
        if (node == null || node.Children.Count == 0) return;
        node.IsExpanded = !node.IsExpanded;
        RefreshRows();
        SelectById(id);
    }

    public void ExpandSelected()
    {
        if (_selectedIndex < 0 || _root == null) return;
        var row = Rows[_selectedIndex];
        var node = FindNode(_root, row.Id);
        if (node == null || node.Children.Count == 0) return;
        if (!node.IsExpanded)
        {
            node.IsExpanded = true;
            RefreshRows();
            SelectById(row.Id);
            return;
        }

        if (_selectedIndex + 1 < Rows.Count && Rows[_selectedIndex + 1].Indent > row.Indent)
        {
            SelectRow(_selectedIndex + 1);
        }
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

    public void HandleCtrlQ()
    {
        if (ExplorerKeyboard.ResolveCtrlQ(_selectedIndex) == ExplorerCtrlQAction.CycleExpand)
        {
            CycleExpandCollapse();
            return;
        }

        DrillRequested?.Invoke();
    }

    public void RequestRefresh() => RefreshRequested?.Invoke();

    public void SelectFormulaToken(int charIndex, bool additive)
    {
        var sheet = TraceUtils.ParseWorksheetScopedAddress(RootAddress)?.WorksheetName ?? "";
        var hit = FormulaTokenHits.HitTest(FormulaText, sheet, charIndex);
        if (hit == null || string.IsNullOrWhiteSpace(hit.Address)) return;

        if (!additive)
        {
            foreach (var row in Rows) row.IsSelected = false;
        }

        var firstMatch = -1;
        for (var i = 0; i < Rows.Count; i++)
        {
            if (!TraceUtils.AddressesReferToSameRange(hit.Address, Rows[i].Location)) continue;
            Rows[i].IsSelected = true;
            if (firstMatch < 0) firstMatch = i;
        }

        if (firstMatch < 0) return;
        if (_selectedIndex != firstMatch)
        {
            _selectedIndex = firstMatch;
            OnPropertyChanged(nameof(SelectedIndex));
        }
        else
        {
            OnPropertyChanged(nameof(SelectedIndex));
        }

        NavigateRequested?.Invoke(Rows[firstMatch]);
    }

    public IReadOnlyList<FormulaHighlightSpan> SelectedFormulaSpans()
    {
        var selected = new List<FormulaHighlightSpan>();
        foreach (var row in Rows)
        {
            if (!row.IsSelected) continue;
            selected.Add(new FormulaHighlightSpan
            {
                Start = row.SourceStart,
                Length = row.SourceLength,
                IsOrigin = row.Kind == FormulaNodeKind.Root
            });
        }

        return FormulaHighlight.ForSelectedRows(selected);
    }

    public IReadOnlyList<string> SelectedLocations()
    {
        var locations = new List<string>();
        foreach (var row in Rows)
        {
            if (row.IsSelected && !string.IsNullOrWhiteSpace(row.NavigationAddress))
            {
                locations.Add(row.NavigationAddress);
            }
        }

        return locations;
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

    public void BeginFormulaEdit(string? typedPrefix = null)
    {
        if (IsEditingFormula) return;
        FormulaEditSeed = typedPrefix;
        IsEditingFormula = true;
        OnPropertyChanged(nameof(FormulaEditSeed));
    }

    public void CancelFormulaEdit()
    {
        if (!IsEditingFormula) return;
        FormulaEditSeed = null;
        IsEditingFormula = false;
    }

    public void ApplyFormulaEdit(string formula) => FormulaApplyRequested?.Invoke(formula ?? "");

    public void RequestExplorePrecedents() => ExplorePrecedentsRequested?.Invoke();

    public void RequestExploreDependents() => ExploreDependentsRequested?.Invoke();

    public void RequestGoToExcel() => GoToExcelRequested?.Invoke();

    public void ReselectCurrent()
    {
        if (_selectedIndex < 0 || _selectedIndex >= Rows.Count) return;
        NavigateRequested?.Invoke(Rows[_selectedIndex]);
    }

    public string SelectedNavigationAddress()
    {
        if (_selectedIndex >= 0 && _selectedIndex < Rows.Count &&
            !string.IsNullOrWhiteSpace(Rows[_selectedIndex].NavigationAddress))
        {
            return Rows[_selectedIndex].NavigationAddress;
        }

        return RootAddress;
    }

    public string SelectedAnchorAddress()
    {
        var address = SelectedNavigationAddress();
        if (string.IsNullOrWhiteSpace(address)) return RootAddress;
        var anchor = TraceUtils.VisibleAnchor(address);
        if (string.IsNullOrEmpty(anchor)) return address;
        var parsed = TraceUtils.ParseWorksheetScopedAddress(address);
        if (parsed == null) return TraceUtils.QualifyAddress(anchor, RootAddress);
        return $"'{parsed.WorksheetName}'!{anchor}";
    }

    private void AppendNode(FormulaAstNode node, int depth, bool[] ancestorContinues, bool isLast)
    {
        var guides = new System.Collections.Generic.List<ExplorerTreeGuide>(depth);
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

        var navigation = ExplorerLocations.NavigationAddress(node, RootAddress);
        var location = string.IsNullOrEmpty(navigation)
            ? ""
            : TraceUtils.FormatExplorerLocation(navigation, RootAddress);
        var component = node.Label;
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
            Indent = depth,
            Guides = guides,
            Component = component,
            Info = node.Info,
            Value = ExplorerLocations.DisplayValue(node, _root),
            Location = location,
            NavigationAddress = navigation,
            IsActiveBranch = node.IsActiveBranch,
            HasChildren = node.Children.Count > 0,
            IsExpanded = node.IsExpanded,
            IsSelected = false,
            SourceStart = node.SourceStart,
            SourceLength = node.SourceLength,
            Kind = node.Kind
        });

        if (!node.IsExpanded || node.Children.Count == 0) return;

        var childContinues = new bool[depth + 1];
        for (var level = 0; level < depth; level++) childContinues[level] = ancestorContinues[level];
        childContinues[depth] = !isLast;
        for (var i = 0; i < node.Children.Count; i++)
        {
            AppendNode(node.Children[i], depth + 1, childContinues, i == node.Children.Count - 1);
        }
    }

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
