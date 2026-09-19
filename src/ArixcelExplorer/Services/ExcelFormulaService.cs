using System;
using ArixcelExplorer.Core.Formulas;
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
            if (parsed == null) return false;
            var sheet = FindWorksheet(parsed.WorksheetName);
            if (sheet == null) return false;
            var range = sheet.Range[parsed.RangeAddress];
            value = range.Value2?.ToString() ?? "";
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
        var value = cell.Value2?.ToString() ?? "";
        var tree = FormulaAstParser.Parse(formula, address, value);
        if (ws != null)
        {
            FormulaEvaluator.EnrichTree(tree, _evalContext, ws.Name);
        }

        return tree;
    }
}
