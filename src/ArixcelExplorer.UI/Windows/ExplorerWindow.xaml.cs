using System.Windows;
using System.Windows.Input;
using ArixcelExplorer.UI.ViewModels;

namespace ArixcelExplorer.UI.Windows;

public partial class ExplorerWindow : Window
{
    private readonly ExplorerViewModel _viewModel;

    public ExplorerWindow(ExplorerViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = _viewModel;
        TreeGrid.ItemsSource = _viewModel.Rows;
        _viewModel.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(ExplorerViewModel.FormulaText))
            {
                FormulaBox.Text = _viewModel.FormulaText;
            }

            if (args.PropertyName == nameof(ExplorerViewModel.RootAddress))
            {
                RootAddressText.Text = _viewModel.RootAddress;
            }

            if (args.PropertyName == nameof(ExplorerViewModel.StatusText))
            {
                StatusText.Text = _viewModel.StatusText;
            }
        };
        FormulaBox.Text = _viewModel.FormulaText;
        RootAddressText.Text = _viewModel.RootAddress;
        StatusText.Text = _viewModel.StatusText;
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
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
            case Key.Right:
                _viewModel.ToggleExpandSelected();
                e.Handled = true;
                break;
            case Key.Left:
                _viewModel.ToggleExpandSelected();
                e.Handled = true;
                break;
            case Key.Q when Keyboard.Modifiers == ModifierKeys.Control:
                _viewModel.ExpandAll();
                e.Handled = true;
                break;
            case Key.Enter:
                _viewModel.RequestClose();
                e.Handled = true;
                break;
            case Key.Escape:
                _viewModel.RequestClose();
                e.Handled = true;
                break;
            case Key.Q when Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift):
                _viewModel.RequestDrillDown();
                e.Handled = true;
                break;
        }
    }

    private void TreeGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (TreeGrid.SelectedIndex >= 0)
        {
            _viewModel.SelectRow(TreeGrid.SelectedIndex);
        }
    }
}
