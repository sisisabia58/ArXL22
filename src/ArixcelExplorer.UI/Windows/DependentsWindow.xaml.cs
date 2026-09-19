using System.Windows;
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
        _viewModel.SelectRow(index);
        if (DependentsGrid.SelectedItem != null)
        {
            DependentsGrid.ScrollIntoView(DependentsGrid.SelectedItem);
        }
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
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
}
