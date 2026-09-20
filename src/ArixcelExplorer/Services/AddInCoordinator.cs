using System;
using System.Linq;
using System.Windows;
using System.Windows.Forms.Integration;
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
    private static HighlightService? _highlightService;
    private static readonly ExplorerSession Session = new();
    private static ArixcelOptions _options = ArixcelOptions.Default;
    private static bool _sessionWired;

    public static void Initialize(Excel.Application application)
    {
        _app = application;
        _traceService = new ExcelTraceService(application);
        _formulaService = new ExcelFormulaService(application);
        _auditService = new ExcelAuditService(application);
        _compareService = new ExcelCompareService(application);
        _utilityService = new UtilityShortcutService(application);
        _highlightService = new HighlightService(application);
        EnsureWpfApp();
        EnsureSessionWired();
    }

    public static ArixcelOptions Options => _options;

    public static void OpenExplorer()
    {
        EnsureInitialized();
        try
        {
            var origin = CaptureActiveOrigin();
            ShowExplorerForActiveCell(origin);
        }
        catch (Exception ex)
        {
            ShowCallError("OpenExplorer", ex);
        }
    }

    public static void OpenDependents()
    {
        EnsureInitialized();
        try
        {
            OpenDependentsCore();
        }
        catch (Exception ex)
        {
            ShowCallError("OpenDependents", ex);
        }
    }

    private static void OpenDependentsCore()
    {
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

        var trace = _traceService.Trace(
            selected,
            TraceDirection.Dependents,
            1,
            _options.TraceSafetyLimit);
        var entries = DependentsAggregator.Aggregate(trace.Rows, TraceDirection.Dependents);

        if (_options.ConfirmLargeDependentScan && entries.Count > _options.MaxDependentsBeforeWarning)
        {
            var confirm = MessageBox.Show(
                $"Found {entries.Count} dependents. Continue?",
                "Arixcel Explorer",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;
        }

        var origin = selected.Count > 0
            ? ToOrigin(selected[0])
            : CaptureActiveOrigin();
        var ownerId = Guid.NewGuid().ToString("N");
        var vm = new DependentsViewModel();
        DependentsWindow? window = null;
        vm.NavigateRequested += row =>
        {
            try
            {
                if (string.IsNullOrWhiteSpace(row.Address)) return;
                _traceService.NavigateToAddress(row.Address, stealFocus: false);
            }
            catch
            {
                // Navigation must not close or hide the dependents window.
            }

            window?.Dispatcher.BeginInvoke(new Action(() =>
            {
                window.Activate();
                window.RestoreKeyboardFocus();
            }), System.Windows.Threading.DispatcherPriority.Input);
        };
        var originValue = selected.Count == 1
            ? TraceUtils.FormatTraceValue(selected[0].Value)
            : "";
        vm.Load(entries, origin.OriginAddress, originValue);

        window = new DependentsWindow(vm, _options.CloseBehavior);
        Session.Track(window, origin, ownerId);
        ShowModeless(window);
        try
        {
            _highlightService!.Apply(ownerId, origin.OriginAddress, _options.OriginHighlight);
            _highlightService.ApplyMany(ownerId, entries.Select(entry => entry.Address), _options.DependentHighlight);
        }
        catch
        {
            // Highlights are optional; the window must still stay visible.
        }
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

    public static void CloseAllExplorers()
    {
        Session.CloseAll();
        _highlightService?.RestoreAll();
    }

    public static void ReturnToOrigin()
    {
        EnsureInitialized();
        if (Session.HasOpenWindows)
        {
            Session.ActivateLatest();
            return;
        }

        if (Session.FirstOrigin != null)
        {
            _traceService!.NavigateToAddress(Session.FirstOrigin.OriginAddress);
        }
    }

    public static void ExecuteUtilityShortcut(string shortcutId) =>
        _utilityService?.Execute(shortcutId);

    private static void ShowExplorerForActiveCell(ExplorerStackEntry origin)
    {
        var cell = _app!.ActiveCell;
        var formula = cell.HasFormula ? cell.Formula?.ToString() ?? "" : "";
        var tree = _formulaService!.BuildExplorerTree(cell);
        var ownerId = Guid.NewGuid().ToString("N");

        var vm = new ExplorerViewModel
        {
            RootAddress = origin.OriginAddress,
            StatusText = "Explorer ready"
        };
        ExplorerWindow? window = null;
        vm.NavigateRequested += row =>
        {
            try
            {
                if (string.IsNullOrWhiteSpace(row.Location)) return;
                _traceService!.NavigateToAddress(row.Location, stealFocus: false);
                _highlightService!.SetTransient(ownerId, row.Location, _options.PrecedentHighlight, origin.OriginAddress);
            }
            catch
            {
                // Navigation/highlights must not close or hide Explorer.
            }

            window?.Dispatcher.BeginInvoke(new Action(() =>
            {
                window.Activate();
                window.RestoreKeyboardFocus();
            }), System.Windows.Threading.DispatcherPriority.Input);
        };
        vm.LoadTree(tree, formula);

        window = new ExplorerWindow(vm, _options.CloseBehavior);
        Session.Track(window, origin, ownerId);
        ShowModeless(window);
        try
        {
            _highlightService!.Apply(ownerId, origin.OriginAddress, _options.OriginHighlight);
        }
        catch
        {
            // Highlights are optional; the window must still stay visible.
        }
    }

    private static void EnsureSessionWired()
    {
        if (_sessionWired) return;
        Session.WindowClosed += (sessionWindow, mode, previous) =>
        {
            _highlightService?.ReleaseOwner(sessionWindow.HighlightOwnerId);
            if (mode == ExplorerCloseMode.RestorePrevious && previous != null)
            {
                _traceService?.NavigateToAddress(previous.OriginAddress);
            }
        };
        _sessionWired = true;
    }

    private static ExplorerStackEntry CaptureActiveOrigin()
    {
        var cell = _app!.ActiveCell;
        var ws = cell.Worksheet as Excel.Worksheet;
        return new ExplorerStackEntry
        {
            OriginAddress = $"'{ws?.Name}'!{cell.Address[false, false]}",
            WorksheetName = ws?.Name ?? "",
            RowIndex = cell.Row - 1,
            ColumnIndex = cell.Column - 1
        };
    }

    private static ExplorerStackEntry ToOrigin(TraceCellInfo cell) => new()
    {
        OriginAddress = cell.Address,
        WorksheetName = cell.WorksheetName,
        RowIndex = cell.RowIndex,
        ColumnIndex = cell.ColumnIndex
    };

    private static void ShowModeless(Window window)
    {
        EnsureWpfApp();
        window.ShowInTaskbar = true;
        window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        window.Topmost = true;
        try
        {
            ElementHost.EnableModelessKeyboardInterop(window);
        }
        catch
        {
            // keyboard interop is best-effort on hosts without WinForms
        }

        window.Show();
        window.WindowState = WindowState.Normal;
        window.Activate();
        window.Focus();
        window.Dispatcher.BeginInvoke(new Action(() =>
        {
            if (!window.IsVisible) return;
            window.Topmost = false;
            window.Activate();
            window.Focus();
        }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
    }

    private static void ShowCallError(string operation, Exception ex)
    {
        var inner = Unwrap(ex);
        MessageBox.Show(
            $"Could not open {operation}: {inner.Message}",
            "Arixcel Explorer",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }

    internal static Exception Unwrap(Exception ex)
    {
        if (ex is AggregateException aggregate)
        {
            return aggregate.Flatten().InnerException ?? aggregate;
        }

        return ex.InnerException ?? ex;
    }

    private static void EnsureWpfApp()
    {
        if (Application.Current == null)
        {
            new Application
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown
            };
        }
    }

    private static void EnsureInitialized()
    {
        if (_app == null || _traceService == null || _highlightService == null)
        {
            throw new InvalidOperationException("Arixcel Explorer is not initialized.");
        }

        EnsureSessionWired();
    }
}
