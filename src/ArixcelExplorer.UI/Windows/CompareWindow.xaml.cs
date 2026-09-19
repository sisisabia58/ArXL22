using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ArixcelExplorer.Core.Compare;

namespace ArixcelExplorer.UI.Windows;

public partial class CompareWindow : Window
{
    public CompareWindow(IReadOnlyList<CompareDiffEntry> diffs)
    {
        InitializeComponent();
        DiffGrid.ItemsSource = diffs;
        SummaryText.Text = $"{diffs.Count} difference(s) found.";
    }
}
