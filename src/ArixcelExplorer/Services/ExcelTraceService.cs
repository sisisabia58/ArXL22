using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ArixcelExplorer.Core.Tracing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.Services;

public sealed class ExcelTraceService
{
    private readonly Excel.Application _app;

    public ExcelTraceService(Excel.Application application)
    {
        _app = application;
    }

    public IReadOnlyList<TraceCellInfo> GetSelectedCells()
    {
        var result = new List<TraceCellInfo>();
        var selection = _app.Selection as Excel.Range;
        if (selection == null) return result;

        foreach (Excel.Range cell in selection.Cells)
        {
            result.Add(ToTraceCellInfo(cell));
        }

        return result;
    }

    public TraceCellInfo GetActiveCell() => ToTraceCellInfo(_app.ActiveCell);

    public async Task<TraceBuilderResult> TraceAsync(
        IEnumerable<TraceCellInfo> roots,
        TraceDirection direction,
        int maxDepth,
        int maxRows)
    {
        var rootList = roots.ToList();
        if (rootList.Count == 0)
        {
            return new TraceBuilderResult { Rows = Array.Empty<TraceRow>() };
        }

        if (rootList.Count == 1)
        {
            return await TraceBuilder.BuildAsync(new TraceBuilderInput
            {
                Root = rootList[0],
                MaxDepth = maxDepth,
                MaxRows = maxRows,
                GetAllNeighbors = cells => Task.FromResult(GetNeighbors(cells, direction))
            }).ConfigureAwait(false);
        }

        var allRows = new List<TraceRow>();
        var truncated = false;
        foreach (var root in rootList)
        {
            var partial = await TraceBuilder.BuildAsync(new TraceBuilderInput
            {
                Root = root,
                MaxDepth = 1,
                MaxRows = maxRows,
                GetAllNeighbors = cells => Task.FromResult(GetNeighbors(cells, direction))
            }).ConfigureAwait(false);

            allRows.AddRange(partial.Rows);
            truncated |= partial.Truncated;
        }

        return new TraceBuilderResult { Rows = allRows, Truncated = truncated };
    }

    public void NavigateToAddress(string address, bool stealFocus = true)
    {
        if (string.IsNullOrWhiteSpace(address)) return;
        var parsed = TraceUtils.ParseWorksheetScopedAddress(address);
        if (parsed == null) return;

        Excel.Worksheet? sheet = null;
        foreach (Excel.Worksheet ws in _app.Worksheets)
        {
            if (string.Equals(ws.Name, parsed.WorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                sheet = ws;
                break;
            }
        }

        if (sheet == null) return;
        var target = sheet.Range[parsed.RangeAddress];
        var previousUpdating = _app.ScreenUpdating;
        try
        {
            _app.ScreenUpdating = false;
            if (!ReferenceEquals(_app.ActiveSheet, sheet))
            {
                sheet.Activate();
            }

            target.Select();
        }
        finally
        {
            _app.ScreenUpdating = previousUpdating;
        }

        if (!stealFocus)
        {
            TryFocusExcelHwndOwner();
        }
    }

    private void TryFocusExcelHwndOwner()
    {
        // No-op here; WPF window restores focus after this returns.
    }

    public IReadOnlyList<DependentEntry> GetDependentsForSelection(int maxDepth, int maxRows)
    {
        var roots = GetSelectedCells();
        var task = TraceAsync(roots, TraceDirection.Dependents, maxDepth, maxRows);
        task.Wait();
        return DependentsAggregator.Aggregate(task.Result.Rows, TraceDirection.Dependents);
    }

    private IReadOnlyList<IReadOnlyList<TraceCellInfo>> GetNeighbors(
        IReadOnlyList<TraceCellInfo> cells,
        TraceDirection direction)
    {
        var results = new List<IReadOnlyList<TraceCellInfo>>();
        foreach (var cell in cells)
        {
            results.Add(GetDirectNeighbors(cell, direction));
        }

        return results;
    }

    private List<TraceCellInfo> GetDirectNeighbors(TraceCellInfo cell, TraceDirection direction)
    {
        var neighbors = new List<TraceCellInfo>();
        var sheet = FindWorksheet(cell.WorksheetName);
        if (sheet == null) return neighbors;

        var source = sheet.Cells[cell.RowIndex + 1, cell.ColumnIndex + 1] as Excel.Range;
        if (source == null) return neighbors;

        Excel.Range? links = direction == TraceDirection.Precedents
            ? source.DirectPrecedents as Excel.Range
            : source.DirectDependents as Excel.Range;
        if (links == null) return neighbors;

        foreach (Excel.Range linked in links.Cells)
        {
            neighbors.Add(ToTraceCellInfo(linked));
        }

        return neighbors;
    }

    private Excel.Worksheet? FindWorksheet(string name)
    {
        foreach (Excel.Worksheet ws in _app.Worksheets)
        {
            if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return ws;
            }
        }

        return null;
    }

    private static TraceCellInfo ToTraceCellInfo(Excel.Range cell)
    {
        var ws = cell.Worksheet as Excel.Worksheet;
        return new TraceCellInfo
        {
            WorksheetName = ws?.Name ?? "",
            RowIndex = cell.Row - 1,
            ColumnIndex = cell.Column - 1,
            Address = $"'{ws?.Name}'!{cell.Address[false, false]}",
            Value = cell.Value2,
            Formula = cell.HasFormula ? cell.Formula : null
        };
    }
}
