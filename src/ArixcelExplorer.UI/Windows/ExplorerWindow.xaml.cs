using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;
using ArixcelExplorer.UI.ViewModels;

namespace ArixcelExplorer.UI.Windows;

public partial class ExplorerWindow : Window
{
    private static readonly SolidColorBrush FormulaLinkBrush = Freeze(Color.FromRgb(0x00, 0x00, 0xFF));
    private static readonly SolidColorBrush FormulaTextBrush = Freeze(Color.FromRgb(0x1A, 0x1A, 0x1A));

    private static SolidColorBrush Freeze(Color color)
    {
        var brush = new SolidColorBrush(color);
        brush.Freeze();
        return brush;
    }

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
            if (args.PropertyName == nameof(ExplorerViewModel.IsEditingFormula))
            {
                ApplyFormulaEditChrome();
            }

            if (args.PropertyName == nameof(ExplorerViewModel.SelectedIndex) ||
                args.PropertyName == nameof(ExplorerViewModel.FormulaText))
            {
                SyncListSelection();
                if (!_viewModel.IsEditingFormula)
                {
                    PaintFormula();
                }

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
        var sheet = TraceUtils.ParseWorksheetScopedAddress(_viewModel.RootAddress)?.WorksheetName ?? "";
        var links = FormulaTokenHits.Spans(text, sheet);
        var paragraph = new Paragraph { Margin = new Thickness(0) };
        var index = 0;
        while (index < text.Length)
        {
            var link = (FormulaTokenSpan?)null;
            foreach (var candidate in links)
            {
                if (index >= candidate.Start && index < candidate.Start + candidate.Length && candidate.Length > 0)
                {
                    link = candidate;
                    break;
                }
            }

            if (link != null)
            {
                var end = Math.Min(text.Length, link.Start + link.Length);
                paragraph.Inlines.Add(new Run(text.Substring(index, end - index))
                {
                    Foreground = FormulaLinkBrush,
                    TextDecorations = TextDecorations.Underline
                });
                index = end;
                continue;
            }

            var next = text.Length;
            foreach (var candidate in links)
            {
                if (candidate.Start > index && candidate.Start < next)
                {
                    next = candidate.Start;
                }
            }

            paragraph.Inlines.Add(new Run(text.Substring(index, next - index))
            {
                Foreground = FormulaTextBrush
            });
            index = next;
        }

        if (paragraph.Inlines.Count == 0)
        {
            paragraph.Inlines.Add(new Run(text) { Foreground = FormulaTextBrush });
        }

        FormulaBox.Document.PagePadding = new Thickness(0);
        FormulaBox.Document.Blocks.Clear();
        FormulaBox.Document.Blocks.Add(paragraph);
    }

    private void ApplyFormulaEditChrome()
    {
        if (_viewModel.IsEditingFormula)
        {
            FormulaBox.Visibility = Visibility.Collapsed;
            FormulaEditor.Visibility = Visibility.Visible;
            var seed = _viewModel.FormulaEditSeed ?? "";
            FormulaEditor.Text = (_viewModel.FormulaText ?? "") + seed;
            FormulaEditor.Focus();
            FormulaEditor.CaretIndex = FormulaEditor.Text.Length;
            return;
        }

        FormulaEditor.Visibility = Visibility.Collapsed;
        FormulaBox.Visibility = Visibility.Visible;
        PaintFormula();
        RestoreKeyboardFocus();
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
        if (_viewModel.IsEditingFormula)
        {
            if (key == Key.Enter)
            {
                _viewModel.ApplyFormulaEdit(FormulaEditor.Text);
                return true;
            }

            if (key == Key.Escape)
            {
                _viewModel.CancelFormulaEdit();
                return true;
            }

            return false;
        }

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
        if (_viewModel.IsEditingFormula)
        {
            if (e.Key == Key.Enter)
            {
                _viewModel.ApplyFormulaEdit(FormulaEditor.Text);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                _viewModel.CancelFormulaEdit();
                e.Handled = true;
            }

            return;
        }

        if (e.Key == Key.System && HandleAltKey(e.SystemKey))
        {
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
            case Key.F2:
                _viewModel.BeginFormulaEdit();
                e.Handled = true;
                break;
            case Key.Back when Keyboard.Modifiers == ModifierKeys.Control:
                _viewModel.RequestGoToExcel();
                e.Handled = true;
                break;
            case Key.Q when Keyboard.Modifiers == ModifierKeys.Control:
                _viewModel.HandleCtrlQ();
                e.Handled = true;
                break;
            case Key.Q when Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift):
                _viewModel.RequestExploreDependents();
                e.Handled = true;
                break;
            case Key.W when Keyboard.Modifiers == ModifierKeys.Control:
                _viewModel.RequestExplorePrecedents();
                e.Handled = true;
                break;
            case Key.L when Keyboard.Modifiers == ModifierKeys.Control:
                _viewModel.RequestExplorePrecedents();
                e.Handled = true;
                break;
            case Key.R when Keyboard.Modifiers == ModifierKeys.Alt:
                _viewModel.RequestRefresh();
                e.Handled = true;
                break;
            case Key.S when Keyboard.Modifiers == ModifierKeys.Alt:
                _viewModel.ReselectCurrent();
                e.Handled = true;
                break;
            case Key.Z when Keyboard.Modifiers == ModifierKeys.Alt:
                _viewModel.CycleExpandCollapse();
                e.Handled = true;
                break;
        }
    }

    private void Window_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (_viewModel.IsEditingFormula) return;
        if (Keyboard.Modifiers != ModifierKeys.None) return;
        if (string.IsNullOrEmpty(e.Text) || char.IsControl(e.Text[0])) return;
        _viewModel.BeginFormulaEdit(e.Text);
        e.Handled = true;
    }

    private bool HandleAltKey(Key key)
    {
        switch (key)
        {
            case Key.R:
                _viewModel.RequestRefresh();
                return true;
            case Key.S:
                _viewModel.ReselectCurrent();
                return true;
            case Key.Z:
                _viewModel.CycleExpandCollapse();
                return true;
            default:
                return false;
        }
    }

    private void FormulaEditor_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            _viewModel.ApplyFormulaEdit(FormulaEditor.Text);
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            _viewModel.CancelFormulaEdit();
            e.Handled = true;
        }
    }

    private void More_Click(object sender, RoutedEventArgs e)
    {
        MoreMenu.PlacementTarget = MoreButton;
        MoreMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        Dispatcher.BeginInvoke(new Action(() => MoreMenu.IsOpen = true));
    }

    private void EditFormula_Click(object sender, RoutedEventArgs e) => _viewModel.BeginFormulaEdit();

    private void ExplorePrecedents_Click(object sender, RoutedEventArgs e) => _viewModel.RequestExplorePrecedents();

    private void ExploreDependents_Click(object sender, RoutedEventArgs e) => _viewModel.RequestExploreDependents();

    private void GoToExcel_Click(object sender, RoutedEventArgs e) => _viewModel.RequestGoToExcel();

    private void Refresh_Click(object sender, RoutedEventArgs e) => _viewModel.RequestRefresh();

    private void Reselect_Click(object sender, RoutedEventArgs e) => _viewModel.ReselectCurrent();

    private void ExpandCollapse_Click(object sender, RoutedEventArgs e) => _viewModel.CycleExpandCollapse();

    private void ExploreCurrentLocation_Click(object sender, RoutedEventArgs e) => _viewModel.RequestExplorePrecedents();

    private void ResetColumns_Click(object sender, RoutedEventArgs e)
    {
        if (TreeGrid.View is not GridView view || view.Columns.Count < 4) return;
        view.Columns[0].Width = ExplorerChrome.ElementWidth;
        view.Columns[1].Width = ExplorerChrome.InfoWidth;
        view.Columns[2].Width = ExplorerChrome.ValueWidth;
        view.Columns[3].Width = ExplorerChrome.LocationWidth;
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
