using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Settings;
using ArixcelExplorer.Core.Tracing;
using ArixcelExplorer.UI.ViewModels;
using ArixcelExplorer.UI.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.Services;

public static class AddInCoordinator
{
    private static Excel.Application? _app;
    private static ExcelTraceService? _traceService;
    private static ExcelFormulaService? _formulaService;
    private static ExcelAuditService? _auditService;
    private static ExcelCompareService? _compareService;
    private static UtilityShortcutService? _utilityService;
    private static readonly ExplorerStack ExplorerStack = new();
    private static ArixcelOptions _options = ArixcelOptions.Default;

    public static void Initialize(Excel.Application application)
    {
        _app = application;
        _traceService = new ExcelTraceService(application);
        _formulaService = new ExcelFormulaService(application);
        _auditService = new ExcelAuditService(application);
        _compareService = new ExcelCompareService(application);
        _utilityService = new UtilityShortcutService(application);
    }

    public static ArixcelOptions Options => _options;

    public static void OpenExplorer()
    {
        EnsureInitialized();
        var cell = _app!.ActiveCell;
        var ws = cell.Worksheet as Excel.Worksheet;
        var address = $"'{ws?.Name}'!{cell.Address[false, false]}";
        ExplorerStack.Push(new ExplorerStackEntry
        {
            OriginAddress = address,
            WorksheetName = ws?.Name ?? "",
            RowIndex = cell.Row - 1,
            ColumnIndex = cell.Column - 1
        });

        ShowExplorerForActiveCell();
    }

    public static void OpenDependents()
    {
        EnsureInitialized();
        var selected = _traceService!.GetSelectedCells();
        if (selected.Count == 0) return;

        if (_options.ConfirmLargeDependentScan && selected.Count > 1)
        {
            var confirm = MessageBox.Show(
                $"Scan dependents for {selected.Count} selected cells? Large scans may be slow.",
                "Arixcel Explorer",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;
        }

        var task = _traceService.TraceAsync(
            selected,
            TraceDirection.Dependents,
            1,
            _options.TraceSafetyLimit);
        task.Wait();
        var entries = DependentsAggregator.Aggregate(task.Result.Rows, TraceDirection.Dependents);

        if (_options.ConfirmLargeDependentScan && entries.Count > _options.MaxDependentsBeforeWarning)
        {
            var confirm = MessageBox.Show(
                $"Found {entries.Count} dependents. Continue?",
                "Arixcel Explorer",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;
        }

        var vm = new DependentsViewModel
        {
            SourceSummary = selected.Count == 1
                ? selected[0].Address
                : $"{selected.Count} selected cells"
        };
        vm.Load(entries, vm.SourceSummary);
        vm.NavigateRequested += row => _traceService.NavigateToAddress(row.Address);
        vm.CloseRequested += () => { /* window closes via dialog result */ };

        var window = new DependentsWindow(vm);
        window.ShowDialog();
    }

    public static void OpenFormulaMap()
    {
        EnsureInitialized();
        _auditService!.ApplyFormulaMap(_options);
    }

    public static void OpenCalculationFlow()
    {
        EnsureInitialized();
        _auditService!.ApplyCalculationFlow();
    }

    public static void OpenCompare()
    {
        EnsureInitialized();
        var active = (_app!.ActiveSheet as Excel.Worksheet)?.Name ?? "";
        var names = _app.Worksheets.Cast<Excel.Worksheet>()
            .Select(ws => ws.Name)
            .Where(n => !string.Equals(n, active, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (names.Count == 0) return;

        var diffs = _compareService!.CompareActiveSheetTo(names[0]);
        var window = new CompareWindow(diffs);
        window.ShowDialog();
    }

    public static void OpenOptions()
    {
        var window = new OptionsWindow(_options);
        if (window.ShowDialog() == true)
        {
            _options = window.Options;
        }
    }

    public static void ClearFormulaMap() => _auditService?.ClearOverlays();

    public static void ExecuteUtilityShortcut(string shortcutId) =>
        _utilityService?.Execute(shortcutId);

    private static void ShowExplorerForActiveCell()
    {
        var cell = _app!.ActiveCell;
        var ws = cell.Worksheet as Excel.Worksheet;
        var address = $"'{ws?.Name}'!{cell.Address[false, false]}";
        var formula = cell.HasFormula ? cell.Formula?.ToString() ?? "" : "";
        var tree = _formulaService!.BuildExplorerTree(cell);

        var vm = new ExplorerViewModel
        {
            RootAddress = address,
            FormulaText = formula,
            StatusText = "Explorer ready"
        };
        vm.LoadTree(tree, formula);
        vm.NavigateRequested += row =>
        {
            if (!string.IsNullOrWhiteSpace(row.Location))
            {
                _traceService!.NavigateToAddress(row.Location);
            }
        };
        vm.DrillDownRequested += () =>
        {
            OpenExplorer();
        };
        vm.CloseRequested += () =>
        {
            if (_options.CloseBehavior == ExplorerCloseBehavior.EscNavigatesBack)
            {
                var previous = ExplorerStack.Pop();
                if (previous != null)
                {
                    _traceService!.NavigateToAddress(previous.OriginAddress);
                }
            }
        };

        var window = new ExplorerWindow(vm);
        window.ShowDialog();
    }

    private static void EnsureInitialized()
    {
        if (_app == null || _traceService == null)
        {
            throw new InvalidOperationException("Arixcel Explorer is not initialized.");
        }
    }
}
