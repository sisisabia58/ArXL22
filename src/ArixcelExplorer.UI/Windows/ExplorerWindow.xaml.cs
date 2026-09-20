using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.UI.ViewModels;

namespace ArixcelExplorer.UI.Windows;

public partial class ExplorerWindow : Window
{
    private readonly ExplorerViewModel _viewModel;
    private readonly ExplorerCloseBehavior _closeBehavior;
    private bool _explicitClose;
    private bool _syncingSelection;

    public ExplorerCloseMode CloseMode { get; private set; } = ExplorerCloseMode.KeepSelection;

    public ExplorerWindow(ExplorerViewModel viewModel, ExplorerCloseBehavior closeBehavior)
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
            if (args.PropertyName == nameof(ExplorerViewModel.SelectedIndex))
            {
                SyncListSelection();
                if (TreeGrid.SelectedItem != null)
                {
                    TreeGrid.ScrollIntoView(TreeGrid.SelectedItem);
                }
            }
        };
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        TreeGrid.Focus();
        if (_viewModel.Rows.Count == 0) return;
        var index = _viewModel.SelectedIndex >= 0 ? _viewModel.SelectedIndex : 0;
        TreeGrid.SelectedIndex = index;
        _viewModel.SelectRow(index);
        SyncListSelection();
        if (TreeGrid.SelectedItem != null)
        {
            TreeGrid.ScrollIntoView(TreeGrid.SelectedItem);
        }
    }

    public void RestoreKeyboardFocus()
    {
        Activate();
        TreeGrid.Focus();
        if (TreeGrid.SelectedItem != null)
        {
            var row = TreeGrid.ItemContainerGenerator.ContainerFromItem(TreeGrid.SelectedItem) as ListViewItem;
            row?.Focus();
            TreeGrid.Focus();
        }
    }

    private void TreeGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_syncingSelection) return;
        if (TreeGrid.SelectedIndex >= 0)
        {
            _viewModel.SelectRow(TreeGrid.SelectedIndex);
        }
    }

    private void FormulaBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var index = FormulaBox.GetCharacterIndexFromPoint(e.GetPosition(FormulaBox), true);
        var additive = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
        _viewModel.SelectFormulaToken(index, additive);
        SyncListSelection();
        RestoreKeyboardFocus();
        e.Handled = true;
    }

    private void SyncListSelection()
    {
        _syncingSelection = true;
        try
        {
            TreeGrid.SelectedItems.Clear();
            foreach (var row in _viewModel.Rows.Where(item => item.IsSelected))
            {
                TreeGrid.SelectedItems.Add(row);
            }
        }
        finally
        {
            _syncingSelection = false;
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
                _viewModel.MoveSelection(-1);
                e.Handled = true;
                break;
            case Key.Down:
                _viewModel.MoveSelection(1);
                e.Handled = true;
                break;
            case Key.Right:
                _viewModel.ExpandSelected();
                e.Handled = true;
                break;
            case Key.Left:
                _viewModel.CollapseSelectedOrMoveToParent();
                e.Handled = true;
                break;
            case Key.Q when Keyboard.Modifiers == ModifierKeys.Control:
                _viewModel.HandleCtrlQ();
                e.Handled = true;
                break;
            case Key.R when Keyboard.Modifiers == ModifierKeys.Alt:
                _viewModel.RequestRefresh();
                e.Handled = true;
                break;
            case Key.Enter:
                _viewModel.RequestKeepClose();
                e.Handled = true;
                break;
            case Key.Escape:
                _viewModel.RequestBackClose();
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
