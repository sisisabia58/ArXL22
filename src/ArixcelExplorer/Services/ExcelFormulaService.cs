using System;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Tracing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.Services;

public sealed class ExcelFormulaEvaluationContext : IFormulaEvaluationContext
{
    private readonly Excel.Application _app;

    public ExcelFormulaEvaluationContext(Excel.Application application)
    {
        _app = application;
    }

    public string? EvaluateSubExpression(string subFormula, string contextWorksheet)
    {
        try
        {
            var sheet = FindWorksheet(contextWorksheet);
            if (sheet == null) return null;
            var result = _app.Evaluate(subFormula);
            return result?.ToString();
        }
        catch
        {
            return null;
        }
    }

    public bool TryResolveReferenceValue(string reference, out string value)
    {
        value = "";
        try
        {
            var parsed = Core.Tracing.TraceUtils.ParseWorksheetScopedAddress(reference);
            if (parsed == null)
            {
                var sheetName = (_app.ActiveSheet as Excel.Worksheet)?.Name ?? "";
                parsed = Core.Tracing.TraceUtils.ParseWorksheetScopedAddress(
                    TraceUtils.QualifyAddress(reference, $"'{sheetName}'!A1"));
            }
            if (parsed == null) return false;
            var sheet = FindWorksheet(parsed.WorksheetName);
            if (sheet == null) return false;
            var range = sheet.Range[parsed.RangeAddress];
            value = TraceUtils.FormatTraceValue(range.Value2);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public int ResolveActiveIfBranch(string formula, System.Collections.Generic.IReadOnlyList<string> branches)
    {
        _ = formula;
        return branches.Count > 1 ? 1 : 0;
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
}

public sealed class ExcelFormulaService
{
    private readonly Excel.Application _app;
    private readonly ExcelFormulaEvaluationContext _evalContext;

    public ExcelFormulaService(Excel.Application application)
    {
        _app = application;
        _evalContext = new ExcelFormulaEvaluationContext(application);
    }

    public FormulaAstNode BuildExplorerTree(Excel.Range cell)
    {
        var ws = cell.Worksheet as Excel.Worksheet;
        var address = $"'{ws?.Name}'!{cell.Address[false, false]}";
        var formula = cell.HasFormula ? cell.Formula?.ToString() ?? "" : "";
        var value = TraceUtils.FormatTraceValue(cell.Value2);
        var tree = FormulaAstParser.Parse(formula, address, value);
        if (ws != null)
        {
            FormulaEvaluator.EnrichTree(tree, _evalContext, ws.Name);
        }

        AttachValidationSource(tree, cell, ws?.Name ?? "");
        return tree;
    }

    private void AttachValidationSource(FormulaAstNode tree, Excel.Range cell, string worksheetName)
    {
        try
        {
            var validation = cell.Validation;
            if (validation == null) return;
            if (Convert.ToInt32(validation.Type) != 3) return;

            var formula1 = validation.Formula1?.ToString() ?? "";
            var source = ValidationListParser.TryParse(formula1, worksheetName);
            if (source == null) return;

            if (source.IsNamedRange)
            {
                var resolved = ResolveNamedRangeAddress(source.Location, worksheetName);
                if (!string.IsNullOrWhiteSpace(resolved))
                {
                    source.Location = resolved!;
                }
            }
            else
            {
                source.Location = TraceUtils.QualifyAddress(source.Location, $"'{worksheetName}'!A1");
            }

            FormulaAstParser.AttachValidationSource(tree, source);
        }
        catch
        {
            // Excel throws when the cell has no validation.
        }
    }

    private string? ResolveNamedRangeAddress(string name, string worksheetName)
    {
        try
        {
            var workbook = _app.ActiveWorkbook;
            if (workbook != null)
            {
                foreach (Excel.Name named in workbook.Names)
                {
                    if (string.Equals(named.Name, name, StringComparison.OrdinalIgnoreCase) ||
                        named.Name.EndsWith("!" + name, StringComparison.OrdinalIgnoreCase))
                    {
                        var range = named.RefersToRange;
                        var ws = range.Worksheet as Excel.Worksheet;
                        return $"'{ws?.Name}'!{range.Address[false, false]}";
                    }
                }
            }
        }
        catch
        {
            // named range may not resolve
        }

        _ = worksheetName;
        return null;
    }
}
