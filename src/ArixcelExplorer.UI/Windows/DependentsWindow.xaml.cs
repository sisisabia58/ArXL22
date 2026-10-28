using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;
using ArixcelExplorer.UI.Services;
using ArixcelExplorer.UI.ViewModels;

namespace ArixcelExplorer.UI.Windows;

public partial class DependentsWindow : Window
{
    private readonly DependentsViewModel _viewModel;
    private readonly ExplorerCloseBehavior _closeBehavior;
    private readonly ExcelSheetHighlight _sheetHighlight = new ExcelSheetHighlight();
    private bool _explicitClose;

    public ExplorerCloseMode CloseMode { get; private set; } = ExplorerCloseMode.KeepSelection;

    public IntPtr WindowHandle => ExplorerWindowFocus.HandleOf(this);

    public DependentsWindow(DependentsViewModel viewModel, ExplorerCloseBehavior closeBehavior)
    {
        InitializeComponent();
        _viewModel = viewModel;
        _closeBehavior = closeBehavior;
        DataContext = _viewModel;
        _viewModel.CloseRequested += mode =>
        {
            CloseMode = mode;
            _explicitClose = true;
            Close();
        };
        Closed += (_, _) => _sheetHighlight.CloseHighlight();
        _viewModel.NavigateRequested += row =>
        {
            if (row == null || row.IsGroup || string.IsNullOrWhiteSpace(row.Address))
            {
                _sheetHighlight.Clear();
                return;
            }

            _sheetHighlight.Show(row.Address, ExcelSheetHighlightColors.Dependent);
        };
        if (DependentsGrid.View is GridView dependentsColumns)
        {
            ExplorerColumnLayout.Attach(
                dependentsColumns,
                "dependents",
                ExplorerChrome.DependentsRangeWidth,
                ExplorerChrome.DependentsCountWidth);
        }
        _viewModel.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(DependentsViewModel.SelectedIndex) &&
                DependentsGrid.SelectedItem != null)
            {
                DependentsGrid.ScrollIntoView(DependentsGrid.SelectedItem);
            }
        };
        ProbeAndApply();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        DependentsGrid.Focus();
        if (_viewModel.Rows.Count == 0) return;
        var index = _viewModel.SelectedIndex >= 0 ? _viewModel.SelectedIndex : 0;
        DependentsGrid.SelectedIndex = index;
        _viewModel.SelectRow(index, navigate: false);
        RestoreKeyboardFocus();
        if (DependentsGrid.SelectedItem != null)
        {
            DependentsGrid.ScrollIntoView(DependentsGrid.SelectedItem);
        }
    }

    public void RestoreKeyboardFocus()
    {
        ExplorerWindowFocus.Reclaim(this);
        DependentsGrid.Focus();
        if (DependentsGrid.SelectedItem != null)
        {
            var row = DependentsGrid.ItemContainerGenerator.ContainerFromItem(DependentsGrid.SelectedItem) as ListViewItem;
            row?.Focus();
            DependentsGrid.Focus();
        }
    }

    public bool TryHandleExplorerKey(Key key)
    {
        switch (key)
        {
            case Key.Up:
                _viewModel.MoveSelection(-1);
                return true;
            case Key.Down:
                _viewModel.MoveSelection(1);
                return true;
            case Key.Left:
                _viewModel.ExpandOrMove(-1);
                return true;
            case Key.Right:
                _viewModel.ExpandOrMove(1);
                return true;
            case Key.Home:
                _viewModel.SelectOrigin();
                return true;
            case Key.Enter:
                _viewModel.RequestKeepClose();
                return true;
            case Key.Escape:
                _viewModel.RequestBackClose();
                return true;
            default:
                return false;
        }
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.System && e.SystemKey == Key.R)
        {
            RefreshAndProbe();
            e.Handled = true;
            return;
        }

        switch (e.Key)
        {
            case Key.Up:
            case Key.Down:
            case Key.Left:
            case Key.Right:
            case Key.Home:
            case Key.Enter:
            case Key.Escape:
                e.Handled = TryHandleExplorerKey(e.Key);
                break;
            case Key.Back when Keyboard.Modifiers == ModifierKeys.Control:
                _viewModel.RequestGoToExcel();
                e.Handled = true;
                break;
            case Key.Q when Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift):
                _viewModel.HandleCtrlShiftQ();
                e.Handled = true;
                break;
            case Key.R when Keyboard.Modifiers == ModifierKeys.Alt:
                RefreshAndProbe();
                e.Handled = true;
                break;
        }
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        _sheetHighlight.CloseHighlight();
        if (_explicitClose) return;
        if (_closeBehavior == ExplorerCloseBehavior.EscNavigatesBack)
        {
            CloseMode = ExplorerCloseMode.RestorePrevious;
        }
    }

    private void Ok_Click(object sender, RoutedEventArgs e) => _viewModel.RequestKeepClose();

    private void More_Click(object sender, RoutedEventArgs e)
    {
        MoreMenu.PlacementTarget = MoreButton;
        MoreMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        Dispatcher.BeginInvoke(new Action(() => MoreMenu.IsOpen = true));
    }

    private void ExploreDependents_Click(object sender, RoutedEventArgs e) => _viewModel.RequestDrill();

    private void GoToExcel_Click(object sender, RoutedEventArgs e) => _viewModel.RequestGoToExcel();

    private void Refresh_Click(object sender, RoutedEventArgs e) => RefreshAndProbe();

    private void TreeGlyph_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is not DependentRow row) return;
        _viewModel.ToggleExpand(row.Key);
        e.Handled = true;
    }

    private void DependentsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        var row = _viewModel.SelectedRow;
        if (row == null || !row.HasChildren) return;
        _viewModel.ToggleExpand(row.Key);
    }

    private void RefreshAndProbe()
    {
        _viewModel.RequestRefresh();
        ProbeAndApply();
    }

    private void ProbeAndApply()
    {
        try
        {
            var excel = ExcelDependentsProbe.TryRead(_viewModel.OriginAddress);
            if (excel == null) return;

            var arrows = new List<string>(excel.Count);
            foreach (var entry in excel)
            {
                arrows.Add(entry.Address);
            }

            var host = new List<string>(_viewModel.Entries.Count);
            foreach (var entry in _viewModel.Entries)
            {
                host.Add(entry.Address);
            }

            var merged = DependentsTree.Combine(_viewModel.OriginAddress, arrows, host);
            _viewModel.ApplyAreas(merged);
        }
        catch
        {
            // Keep the host list when Excel cannot be queried from this window.
        }
    }
}
