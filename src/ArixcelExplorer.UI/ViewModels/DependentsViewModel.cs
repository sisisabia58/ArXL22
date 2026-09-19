using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.UI.ViewModels;

public sealed class DependentRow : INotifyPropertyChanged
{
    public string Address { get; set; } = "";
    public int Count { get; set; }
    public string Value { get; set; } = "";
    public bool IsSelected { get; set; }

    public event PropertyChangedEventHandler? PropertyChanged;
}

public sealed class DependentsViewModel : INotifyPropertyChanged
{
    public ObservableCollection<DependentRow> Rows { get; } = new();
    public string SourceSummary { get; set; } = "";
    public string StatusText { get; set; } = "";

    private int _selectedIndex = -1;

    public event PropertyChangedEventHandler? PropertyChanged;
    public event System.Action<DependentRow>? NavigateRequested;
    public event System.Action? CloseRequested;

    public void Load(IReadOnlyList<DependentEntry> entries, string sourceSummary)
    {
        Rows.Clear();
        SourceSummary = sourceSummary;
        foreach (var entry in entries)
        {
            Rows.Add(new DependentRow
            {
                Address = entry.Address,
                Count = entry.Count,
                Value = entry.Value
            });
        }

        StatusText = $"{Rows.Count} dependent(s)";
        if (Rows.Count > 0) SelectRow(0);
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

    public void RequestClose() => CloseRequested?.Invoke();

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
