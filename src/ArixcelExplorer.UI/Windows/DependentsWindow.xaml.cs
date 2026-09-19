using System.Windows;
using System.Windows.Input;
using ArixcelExplorer.UI.ViewModels;

namespace ArixcelExplorer.UI.Windows;

public partial class DependentsWindow : Window
{
    private readonly DependentsViewModel _viewModel;

    public DependentsWindow(DependentsViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = _viewModel;
        DependentsGrid.ItemsSource = _viewModel.Rows;
        _viewModel.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(DependentsViewModel.SourceSummary))
            {
                SourceSummaryText.Text = _viewModel.SourceSummary;
            }

            if (args.PropertyName == nameof(DependentsViewModel.StatusText))
            {
                StatusText.Text = _viewModel.StatusText;
            }
        };
        SourceSummaryText.Text = _viewModel.SourceSummary;
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
            case Key.Enter:
            case Key.Escape:
                _viewModel.RequestClose();
                e.Handled = true;
                break;
        }
    }

    private void DependentsGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (DependentsGrid.SelectedIndex >= 0)
        {
            _viewModel.SelectRow(DependentsGrid.SelectedIndex);
        }
    }
}
