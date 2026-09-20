using System.Windows;
using ArixcelExplorer.Core.Settings;

namespace ArixcelExplorer.UI.Windows;

public partial class OptionsWindow : Window
{
    public ArixcelOptions Options { get; private set; }

    public OptionsWindow(ArixcelOptions options)
    {
        InitializeComponent();
        Options = options.Clone();
        ConfirmLargeScanBox.IsChecked = Options.ConfirmLargeDependentScan;
        CloseBehaviorCombo.SelectedIndex = Options.CloseBehavior == ExplorerCloseBehavior.EnterKeepsSelection ? 0 : 1;
        OriginHighlightBox.Text = Options.OriginHighlight;
        PrecedentHighlightBox.Text = Options.PrecedentHighlight;
        DependentHighlightBox.Text = Options.DependentHighlight;
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        Options.ConfirmLargeDependentScan = ConfirmLargeScanBox.IsChecked == true;
        Options.CloseBehavior = CloseBehaviorCombo.SelectedIndex == 0
            ? ExplorerCloseBehavior.EnterKeepsSelection
            : ExplorerCloseBehavior.EscNavigatesBack;

        if (ExcelOleColor.IsValidHex(OriginHighlightBox.Text))
        {
            Options.OriginHighlight = NormalizeHex(OriginHighlightBox.Text);
        }

        if (ExcelOleColor.IsValidHex(PrecedentHighlightBox.Text))
        {
            Options.PrecedentHighlight = NormalizeHex(PrecedentHighlightBox.Text);
        }

        if (ExcelOleColor.IsValidHex(DependentHighlightBox.Text))
        {
            Options.DependentHighlight = NormalizeHex(DependentHighlightBox.Text);
        }

        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

    private static string NormalizeHex(string hex)
    {
        var text = hex.Trim();
        return text.StartsWith("#") ? text.ToUpperInvariant() : "#" + text.ToUpperInvariant();
    }
}
