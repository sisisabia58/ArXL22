using System.Windows;
using ArixcelExplorer.Core.Settings;

namespace ArixcelExplorer.UI.Windows;

public partial class OptionsWindow : Window
{
    public ArixcelOptions Options { get; private set; }

    public OptionsWindow(ArixcelOptions options)
    {
        InitializeComponent();
        Options = options;
        MaxDepthBox.Text = options.TraceMaxDepth.ToString();
        DependentsWarningBox.Text = options.MaxDependentsBeforeWarning.ToString();
        ConfirmLargeScanBox.IsChecked = options.ConfirmLargeDependentScan;
        CloseBehaviorCombo.SelectedIndex = options.CloseBehavior == ExplorerCloseBehavior.EnterKeepsSelection ? 0 : 1;
        OriginHighlightBox.Text = options.OriginHighlight;
        PrecedentHighlightBox.Text = options.PrecedentHighlight;
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(MaxDepthBox.Text, out var depth))
        {
            Options.TraceMaxDepth = depth;
        }

        if (int.TryParse(DependentsWarningBox.Text, out var warning))
        {
            Options.MaxDependentsBeforeWarning = warning;
        }

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

        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

    private static string NormalizeHex(string hex)
    {
        var text = hex.Trim();
        return text.StartsWith("#") ? text.ToUpperInvariant() : "#" + text.ToUpperInvariant();
    }
}
