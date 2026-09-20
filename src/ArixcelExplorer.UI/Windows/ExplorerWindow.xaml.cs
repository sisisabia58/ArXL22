using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.UI.ViewModels;

namespace ArixcelExplorer.UI.Windows;

public partial class ExplorerWindow : Window
{
    private static readonly SolidColorBrush FormulaHighlightBrush = new(Color.FromRgb(0x7F, 0xDB, 0xFF));

    private readonly ExplorerViewModel _viewModel;
    private readonly ExplorerCloseBehavior _closeBehavior;
    private bool _explicitClose;
    private bool _syncingSelection;

    public ExplorerCloseMode CloseMode { get; private set; } = ExplorerCloseMode.KeepSelection;

    public IntPtr WindowHandle => ExplorerWindowFocus.HandleOf(this);

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
            if (args.PropertyName == nameof(ExplorerViewModel.SelectedIndex) ||
                args.PropertyName == nameof(ExplorerViewModel.FormulaText))
            {
                SyncListSelection();
                PaintFormula();
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
        if (_viewModel.Rows.Count == 0)
        {
            PaintFormula();
            return;
        }

        var index = _viewModel.SelectedIndex >= 0 ? _viewModel.SelectedIndex : 0;
        TreeGrid.SelectedIndex = index;
        _viewModel.SelectRow(index, navigate: false);
        SyncListSelection();
        PaintFormula();
        if (TreeGrid.SelectedItem != null)
        {
            TreeGrid.ScrollIntoView(TreeGrid.SelectedItem);
        }

        RestoreKeyboardFocus();
    }

    public void RestoreKeyboardFocus()
    {
        ExplorerWindowFocus.Reclaim(this);
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
            PaintFormula();
        }
    }

    private void TreeGlyph_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is not ExplorerTreeRow row) return;
        _viewModel.ToggleExpand(row.Id);
        e.Handled = true;
    }

    private void TreeGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (_viewModel.SelectedIndex < 0 || _viewModel.SelectedIndex >= _viewModel.Rows.Count) return;
        _viewModel.ToggleExpand(_viewModel.Rows[_viewModel.SelectedIndex].Id);
    }

    private void FormulaBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var index = CharIndexFromPoint(FormulaBox, e.GetPosition(FormulaBox));
        var additive = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
        _viewModel.SelectFormulaToken(index, additive);
        SyncListSelection();
        PaintFormula();
        RestoreKeyboardFocus();
        e.Handled = true;
    }

    private void PaintFormula()
    {
        var text = _viewModel.FormulaText ?? "";
        var highlights = _viewModel.SelectedFormulaSpans();
        var paragraph = new Paragraph { Margin = new Thickness(0) };
        var index = 0;
        while (index < text.Length)
        {
            var span = highlights.FirstOrDefault(item => index >= item.Start && index < item.Start + item.Length);
            if (span != null && span.Length > 0)
            {
                var end = Math.Min(text.Length, span.Start + span.Length);
                paragraph.Inlines.Add(new Run(text.Substring(index, end - index))
                {
                    Background = FormulaHighlightBrush
                });
                index = end;
                continue;
            }

            var next = text.Length;
            foreach (var item in highlights)
            {
                if (item.Start > index && item.Start < next)
                {
                    next = item.Start;
                }
            }

            paragraph.Inlines.Add(new Run(text.Substring(index, next - index)));
            index = next;
        }

        if (paragraph.Inlines.Count == 0)
        {
            paragraph.Inlines.Add(new Run(text));
        }

        FormulaBox.Document.Blocks.Clear();
        FormulaBox.Document.Blocks.Add(paragraph);
    }

    private static int CharIndexFromPoint(RichTextBox box, Point point)
    {
        var pointer = box.GetPositionFromPoint(point, true);
        if (pointer == null) return 0;
        var range = new TextRange(box.Document.ContentStart, pointer);
        return range.Text.Replace("\r", "").Replace("\n", "").Length;
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
            case Key.Right:
                _viewModel.ExpandSelected();
                return true;
            case Key.Left:
                _viewModel.CollapseSelectedOrMoveToParent();
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
            case Key.Right:
            case Key.Left:
            case Key.Enter:
            case Key.Escape:
                e.Handled = TryHandleExplorerKey(e.Key);
                break;
            case Key.Q when Keyboard.Modifiers == ModifierKeys.Control:
                _viewModel.HandleCtrlQ();
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
