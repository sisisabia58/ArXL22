using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
    private static ArixcelOptions _options = new();
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
        _options = OptionsStore.Load();
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

    private static void OpenDependentsCore(IReadOnlyList<TraceCellInfo>? roots = null)
    {
        var selected = roots ?? _traceService!.GetSelectedCells();
        if (selected.Count == 0) return;

        if (_options.ConfirmLargeDependentScan && selected.Count > 1)
        {
            var confirm = MessageBox.Show(
                $"Scan dependents for {selected.Count} selected cells? Large scans may be slow.",
                "EXLerate Explorer",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;
        }

        var trace = _traceService!.Trace(
            selected,
            TraceDirection.Dependents,
            1,
            _options.TraceSafetyLimit);
        var entries = DependentsAggregator.Aggregate(trace.Rows, TraceDirection.Dependents);

        if (_options.ConfirmLargeDependentScan && entries.Count > _options.MaxDependentsBeforeWarning)
        {
            var confirm = MessageBox.Show(
                $"Found {entries.Count} dependents. Continue?",
                "EXLerate Explorer",
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
        vm.DrillRequested += () =>
        {
            var row = vm.SelectedRow;
            if (row == null || string.IsNullOrWhiteSpace(row.Address) || row.IsOrigin) return;
            try
            {
                var cell = CellFromAddress(row.Address);
                if (cell != null)
                {
                    OpenDependentsCore(new[] { cell });
                }
            }
            catch (Exception ex)
            {
                ShowCallError("OpenDependents", ex);
            }
        };
        vm.RefreshRequested += () =>
        {
            try
            {
                var cell = CellFromAddress(origin.OriginAddress);
                if (cell == null) return;
                var refreshed = _traceService.Trace(
                    new[] { cell },
                    TraceDirection.Dependents,
                    1,
                    _options.TraceSafetyLimit);
                var refreshedEntries = DependentsAggregator.Aggregate(refreshed.Rows, TraceDirection.Dependents);
                var originValue = TraceUtils.FormatTraceValue(cell.Value);
                vm.Load(refreshedEntries, origin.OriginAddress, originValue);
                _highlightService!.ReleaseOwner(ownerId);
                _highlightService.Apply(ownerId, origin.OriginAddress, _options.OriginHighlight);
                _highlightService.ApplyMany(ownerId, refreshedEntries.Select(entry => entry.Address), _options.DependentHighlight);
            }
            catch
            {
                // Refresh is best-effort.
            }
        };
        var originValueText = selected.Count == 1
            ? TraceUtils.FormatTraceValue(selected[0].Value)
            : "";
        vm.Load(entries, origin.OriginAddress, originValueText);

        window = new DependentsWindow(vm, _options.CloseBehavior);
        Session.Track(window, origin, ownerId);
        ShowModeless(window, _options.DependentsWindow, ExplorerChrome.DependentsWidth, ExplorerChrome.DependentsHeight);
        window.Closed += (_, _) => PersistWindowBounds(window, dependents: true);
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
            window.Options.ExplorerWindow = _options.ExplorerWindow;
            window.Options.DependentsWindow = _options.DependentsWindow;
            _options = window.Options;
            PersistOptions();
        }
    }

    public static void ClearFormulaMap() => _auditService?.ClearOverlays();

    public static void CloseAllExplorers()
    {
        Session.CloseAll();
        _highlightService?.RestoreAll();
        PersistOptions();
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
            RootAddress = origin.OriginAddress
        };
        ExplorerWindow? window = null;
        vm.NavigateRequested += row =>
        {
            try
            {
                if (string.IsNullOrWhiteSpace(row.Location)) return;
                _traceService!.NavigateToAddress(row.Location, stealFocus: false);
                _highlightService!.SetTransientMany(
                    ownerId,
                    vm.SelectedLocations(),
                    _options.PrecedentHighlight,
                    origin.OriginAddress);
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
        vm.DrillRequested += () =>
        {
            try
            {
                ShowExplorerForActiveCell(CaptureActiveOrigin());
            }
            catch (Exception ex)
            {
                ShowCallError("OpenExplorer", ex);
            }
        };
        vm.RefreshRequested += () =>
        {
            try
            {
                var range = RangeFromAddress(origin.OriginAddress);
                if (range == null) return;
                var refreshedFormula = range.HasFormula ? range.Formula?.ToString() ?? "" : "";
                vm.RootAddress = origin.OriginAddress;
                vm.LoadTree(_formulaService.BuildExplorerTree(range), refreshedFormula);
                _highlightService!.Apply(ownerId, origin.OriginAddress, _options.OriginHighlight);
            }
            catch
            {
                // Refresh is best-effort.
            }
        };
        vm.LoadTree(tree, formula);

        window = new ExplorerWindow(vm, _options.CloseBehavior);
        Session.Track(window, origin, ownerId);
        ShowModeless(window, _options.ExplorerWindow, ExplorerChrome.ExplorerWidth, ExplorerChrome.ExplorerHeight);
        window.Closed += (_, _) => PersistWindowBounds(window, dependents: false);
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

    private static void ShowModeless(Window window, WindowPlacement placement, double defaultWidth, double defaultHeight)
    {
        EnsureWpfApp();
        window.ShowInTaskbar = true;
        ApplyPlacement(window, placement, defaultWidth, defaultHeight);
        try
        {
            ElementHost.EnableModelessKeyboardInterop(window);
        }
        catch
        {
            // keyboard interop is best-effort on hosts without WinForms
        }

        window.Topmost = true;
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

    private static void ApplyPlacement(Window window, WindowPlacement placement, double defaultWidth, double defaultHeight)
    {
        window.Width = placement.Width > 0 ? placement.Width : defaultWidth;
        window.Height = placement.Height > 0 ? placement.Height : defaultHeight;
        window.WindowStartupLocation = WindowStartupLocation.Manual;

        if (placement.HasPosition && IsOnScreen(placement))
        {
            window.Left = placement.Left;
            window.Top = placement.Top;
            return;
        }

        if (!TryOffsetFromExcel(window))
        {
            window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }
    }

    private static bool IsOnScreen(WindowPlacement placement)
    {
        var left = SystemParameters.VirtualScreenLeft;
        var top = SystemParameters.VirtualScreenTop;
        var right = left + SystemParameters.VirtualScreenWidth;
        var bottom = top + SystemParameters.VirtualScreenHeight;
        var centerX = placement.Left + placement.Width / 2;
        var centerY = placement.Top + placement.Height / 2;
        return centerX >= left && centerX <= right && centerY >= top && centerY <= bottom;
    }

    private static bool TryOffsetFromExcel(Window window)
    {
        try
        {
            if (_app == null) return false;
            var hwnd = new IntPtr(_app.Hwnd);
            if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out var rect)) return false;

            var dpi = 96u;
            try
            {
                dpi = GetDpiForWindow(hwnd);
            }
            catch
            {
                dpi = 96;
            }

            if (dpi == 0) dpi = 96;
            var scale = 96.0 / dpi;
            window.Left = rect.Left * scale + ExplorerChrome.ExcelOffsetX;
            window.Top = rect.Top * scale + ExplorerChrome.ExcelOffsetY;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void PersistWindowBounds(Window window, bool dependents)
    {
        var placement = new WindowPlacement
        {
            Left = window.Left,
            Top = window.Top,
            Width = window.ActualWidth > 0 ? window.ActualWidth : window.Width,
            Height = window.ActualHeight > 0 ? window.ActualHeight : window.Height
        };

        if (dependents)
        {
            _options.DependentsWindow = placement;
        }
        else
        {
            _options.ExplorerWindow = placement;
        }

        PersistOptions();
    }

    private static void PersistOptions()
    {
        try
        {
            OptionsStore.Save(_options);
        }
        catch
        {
            // Persistence must never block Explorer.
        }
    }

    private static Excel.Range? RangeFromAddress(string address)
    {
        var parsed = TraceUtils.ParseWorksheetScopedAddress(address);
        if (parsed == null || _app == null) return null;
        foreach (Excel.Worksheet ws in _app.Worksheets)
        {
            if (string.Equals(ws.Name, parsed.WorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                return ws.Range[parsed.RangeAddress];
            }
        }

        return null;
    }

    private static TraceCellInfo? CellFromAddress(string address)
    {
        var range = RangeFromAddress(address);
        if (range == null) return null;
        var ws = range.Worksheet as Excel.Worksheet;
        return new TraceCellInfo
        {
            WorksheetName = ws?.Name ?? "",
            RowIndex = range.Row - 1,
            ColumnIndex = range.Column - 1,
            Address = $"'{ws?.Name}'!{range.Address[false, false]}",
            Value = range.Value2,
            Formula = range.HasFormula ? range.Formula : null
        };
    }

    private static void ShowCallError(string operation, Exception ex)
    {
        var inner = Unwrap(ex);
        MessageBox.Show(
            $"Could not open {operation}: {inner.Message}",
            "EXLerate Explorer",
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
            throw new InvalidOperationException("EXLerate Explorer is not initialized.");
        }

        EnsureSessionWired();
    }

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect lpRect);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hWnd);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
