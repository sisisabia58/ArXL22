using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.UI.ViewModels;

public sealed class DependentRow : INotifyPropertyChanged
{
    public string Address { get; set; } = "";
    public int Count { get; set; }
    public string Value { get; set; } = "";
    public bool IsOrigin { get; set; }

    public string AddressDisplay => DependentsAggregator.ElementAddress(Address);
    public string CountDisplay => IsOrigin || Count == 0 ? "" : Count.ToString();

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
    public ObservableCollection<DependentRow> Rows { get; } = new();

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

    private int _selectedIndex = -1;

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
    public event System.Action<DependentRow>? NavigateRequested;
    public event System.Action<ExplorerCloseMode>? CloseRequested;
    public event System.Action? DrillRequested;
    public event System.Action? RefreshRequested;

    public void Load(IReadOnlyList<DependentEntry> entries, string originAddress, string originValue = "")
    {
        Rows.Clear();
        SourceSummary = originAddress;
        var combined = DependentsAggregator.PrependOrigin(originAddress, originValue, entries);
        for (var i = 0; i < combined.Count; i++)
        {
            var entry = combined[i];
            Rows.Add(new DependentRow
            {
                Address = entry.Address,
                Count = entry.Count,
                Value = entry.Value,
                IsOrigin = i == 0
            });
        }

        var dependentCount = System.Math.Max(0, Rows.Count - 1);
        StatusText = $"{dependentCount} dependent(s)";
        if (Rows.Count > 0) SelectRow(0, navigate: false);
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
        else
        {
            _selectedIndex = index;
        }

        if (navigate)
        {
            NavigateRequested?.Invoke(Rows[index]);
        }
    }

    public void MoveSelection(int delta)
    {
        if (Rows.Count == 0) return;
        var next = _selectedIndex < 0 ? 0 : System.Math.Max(0, System.Math.Min(_selectedIndex + delta, Rows.Count - 1));
        SelectRow(next);
    }

    public void RequestKeepClose() => CloseRequested?.Invoke(ExplorerCloseMode.KeepSelection);

    public void RequestBackClose() => CloseRequested?.Invoke(ExplorerCloseMode.RestorePrevious);

    public void RequestRefresh() => RefreshRequested?.Invoke();

    public void HandleCtrlShiftQ()
    {
        if (!ExplorerKeyboard.ShouldDrillDependents(_selectedIndex)) return;
        DrillRequested?.Invoke();
    }

    public DependentRow? SelectedRow =>
        _selectedIndex >= 0 && _selectedIndex < Rows.Count ? Rows[_selectedIndex] : null;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
