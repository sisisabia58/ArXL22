using System.Collections.Generic;
using ArixcelExplorer.Core.Compare;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.Services;

public sealed class ExcelCompareService
{
    private readonly Excel.Application _app;

    public ExcelCompareService(Excel.Application application)
    {
        _app = application;
    }

    public IReadOnlyList<CompareDiffEntry> CompareActiveSheetTo(string otherSheetName)
    {
        var leftSheet = _app.ActiveSheet as Excel.Worksheet;
        Excel.Worksheet? rightSheet = null;
        foreach (Excel.Worksheet ws in _app.Worksheets)
        {
            if (ws.Name == otherSheetName)
            {
                rightSheet = ws;
                break;
            }
        }

        if (leftSheet == null || rightSheet == null)
        {
            return new List<CompareDiffEntry>();
        }

        var leftUsed = leftSheet.UsedRange;
        var rightUsed = rightSheet.UsedRange;
        var leftGrid = SnapshotRange(leftUsed);
        var rightGrid = SnapshotRange(rightUsed);
        return CompareEngine.DiffGrids(leftGrid, rightGrid, leftSheet.Name, rightSheet.Name);
    }

    private static CompareCellSnapshot[][] SnapshotRange(Excel.Range used)
    {
        var rows = used.Rows.Count;
        var cols = used.Columns.Count;
        var grid = new CompareCellSnapshot[rows][];
        for (var r = 0; r < rows; r++)
        {
            grid[r] = new CompareCellSnapshot[cols];
            for (var c = 0; c < cols; c++)
            {
                var cell = used.Cells[r + 1, c + 1] as Excel.Range;
                grid[r][c] = new CompareCellSnapshot
                {
                    Address = cell?.Address[false, false] ?? "",
                    Value = cell?.Value2,
                    Formula = cell?.HasFormula == true ? cell.Formula?.ToString() : null,
                    NumberFormat = cell?.NumberFormat?.ToString()
                };
            }
        }

        return grid;
    }
}
