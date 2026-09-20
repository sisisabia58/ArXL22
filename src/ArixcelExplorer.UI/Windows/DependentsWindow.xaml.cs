using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.UI.ViewModels;

namespace ArixcelExplorer.UI.Windows;

public partial class DependentsWindow : Window
{
    private readonly DependentsViewModel _viewModel;
    private readonly ExplorerCloseBehavior _closeBehavior;
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
        _viewModel.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(DependentsViewModel.SelectedIndex) &&
                DependentsGrid.SelectedItem != null)
            {
                DependentsGrid.ScrollIntoView(DependentsGrid.SelectedItem);
            }
        };
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
            _viewModel.RequestRefresh();
            e.Handled = true;
            return;
        }

        switch (e.Key)
        {
            case Key.Up:
            case Key.Down:
            case Key.Enter:
            case Key.Escape:
                e.Handled = TryHandleExplorerKey(e.Key);
                break;
            case Key.Q when Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift):
                _viewModel.HandleCtrlShiftQ();
                e.Handled = true;
                break;
            case Key.R when Keyboard.Modifiers == ModifierKeys.Alt:
                _viewModel.RequestRefresh();
                e.Handled = true;
                break;
        }
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_explicitClose) return;
        if (_closeBehavior == ExplorerCloseBehavior.EscNavigatesBack)
        {
            CloseMode = ExplorerCloseMode.RestorePrevious;
        }
    }

    private void Ok_Click(object sender, RoutedEventArgs e) => _viewModel.RequestKeepClose();
}
