using System;
using System.Collections.Generic;
using ArixcelExplorer.Core.Audit;
using ArixcelExplorer.Core.Settings;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.Services;

public sealed class ExcelAuditService
{
    private readonly Excel.Application _app;
    private readonly Dictionary<string, int> _originalColors = new(StringComparer.OrdinalIgnoreCase);

    public ExcelAuditService(Excel.Application application)
    {
        _app = application;
    }

    public void ApplyFormulaMap(ArixcelOptions options)
    {
        var selection = _app.Selection as Excel.Range;
        if (selection == null) return;

        var rows = selection.Rows.Count;
        var cols = selection.Columns.Count;
        var cells = new AutoColorCell[rows][];
        var addresses = new string[rows][];
        var consistencyInput = new FormulaConsistencyCell[rows][];

        for (var r = 0; r < rows; r++)
        {
            cells[r] = new AutoColorCell[cols];
            addresses[r] = new string[cols];
            consistencyInput[r] = new FormulaConsistencyCell[cols];
            for (var c = 0; c < cols; c++)
            {
                var cell = selection.Cells[r + 1, c + 1] as Excel.Range;
                var formula = cell?.HasFormula == true ? cell.Formula?.ToString() : null;
                cells[r][c] = new AutoColorCell
                {
                    Formula = formula,
                    Value = cell?.Value2,
                    NumberFormat = cell?.NumberFormat?.ToString(),
                    HasHyperlink = cell?.Hyperlinks.Count > 0
                };
                addresses[r][c] = cell?.Address[false, false] ?? "";
                consistencyInput[r][c] = new FormulaConsistencyCell
                {
                    IsFormula = cell?.HasFormula == true,
                    FormulaR1C1 = cell?.FormulaR1C1?.ToString()
                };
            }
        }

        var consistency = FormulaConsistency.AnalyzeHorizontalFormulaConsistency(consistencyInput);
        var map = FormulaMapService.BuildMap(cells, addresses, consistency, options.MapPalette);
        foreach (var item in map)
        {
            ApplyFill(item.Address, ParseColor(item.FillColor));
        }
    }

    public void ApplyCalculationFlow()
    {
        var selection = _app.Selection as Excel.Range;
        if (selection == null) return;

        var rows = selection.Rows.Count;
        var cols = selection.Columns.Count;
        var grid = new FlowCellInfo[rows][];
        for (var r = 0; r < rows; r++)
        {
            grid[r] = new FlowCellInfo[cols];
            for (var c = 0; c < cols; c++)
            {
                var cell = selection.Cells[r + 1, c + 1] as Excel.Range;
                grid[r][c] = new FlowCellInfo
                {
                    Address = cell?.Address[false, false] ?? "",
                    HasFormula = cell?.HasFormula == true,
                    HasPrecedentsInSelection = cell?.DirectPrecedents != null,
                    HasDependentsInSelection = cell?.DirectDependents != null
                };
            }
        }

        var roles = CalculationFlow.ClassifyGrid(grid);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var cell = selection.Cells[r + 1, c + 1] as Excel.Range;
                ApplyFillToRange(cell, ParseColor(CalculationFlow.GetFlowColor(roles[r][c])));
            }
        }
    }

    public void ClearOverlays()
    {
        foreach (var pair in _originalColors)
        {
            try
            {
                var range = _app.Range[pair.Key];
                range.Interior.Color = pair.Value;
            }
            catch
            {
                // ignore restore failures
            }
        }

        _originalColors.Clear();
    }

    private void ApplyFill(string address, int color)
    {
        var range = _app.Range[address];
        ApplyFillToRange(range, color);
    }

    private void ApplyFillToRange(Excel.Range? range, int color)
    {
        if (range == null) return;
        var key = range.Address[false, false];
        if (!_originalColors.ContainsKey(key))
        {
            _originalColors[key] = (int)range.Interior.Color;
        }

        range.Interior.Color = color;
    }

    private static int ParseColor(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex) || !hex.StartsWith("#")) return 0xFFFFFF;
        var value = Convert.ToInt32(hex.Substring(1), 16);
        var r = (value >> 16) & 0xFF;
        var g = (value >> 8) & 0xFF;
        var b = value & 0xFF;
        return r | (g << 8) | (b << 16);
    }
}
