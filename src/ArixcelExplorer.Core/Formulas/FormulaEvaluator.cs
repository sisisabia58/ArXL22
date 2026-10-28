using System;
using System.Collections.Generic;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.Core.Formulas;

public interface IFormulaEvaluationContext
{
    string? EvaluateSubExpression(string subFormula, string contextWorksheet);
    bool TryResolveReferenceValue(string reference, out string value);
    int ResolveActiveIfBranch(string formula, IReadOnlyList<string> branches);
}

public sealed class NullFormulaEvaluationContext : IFormulaEvaluationContext
{
    public string? EvaluateSubExpression(string subFormula, string contextWorksheet) => null;
    public bool TryResolveReferenceValue(string reference, out string value)
    {
        value = "";
        return false;
    }

    public int ResolveActiveIfBranch(string formula, IReadOnlyList<string> branches) => 0;
}

public static class FormulaEvaluator
{
    public static void EnrichTree(FormulaAstNode root, IFormulaEvaluationContext context, string worksheetName)
    {
        EnrichNode(root, root, context, worksheetName, root.SourceFormula ?? "");
    }

    private static void EnrichNode(
        FormulaAstNode node,
        FormulaAstNode root,
        IFormulaEvaluationContext context,
        string worksheetName,
        string sourceFormula)
    {
        foreach (var child in node.Children)
        {
            EnrichNode(child, root, context, worksheetName, sourceFormula);
        }

        switch (node.Kind)
        {
            case FormulaNodeKind.Reference:
            case FormulaNodeKind.NamedRange:
                var location = TraceUtils.QualifyAddress(
                    string.IsNullOrWhiteSpace(node.Location) ? node.Label : node.Location ?? "",
                    $"'{worksheetName}'!A1");
                if (!string.IsNullOrEmpty(location))
                {
                    node.Location = location;
                }

                if (TraceUtils.IsMultiCellRange(node.Location) || TraceUtils.IsMultiCellRange(node.Label))
                {
                    node.Value = TraceUtils.ArrayValue;
                    break;
                }

                if (node.Location != null && context.TryResolveReferenceValue(node.Location, out var refValue))
                {
                    node.Value = refValue;
                }

                break;
            case FormulaNodeKind.Function:
                EnrichFunction(node, root, context, worksheetName, sourceFormula);
                break;
            case FormulaNodeKind.Operator:
                if (!string.IsNullOrWhiteSpace(node.Label))
                {
                    var evaluated = context.EvaluateSubExpression("=" + node.Label, worksheetName);
                    if (TraceUtils.IsUsableFormulaResult(evaluated))
                    {
                        node.Value = evaluated;
                    }
                }

                break;
        }
    }

    private static void EnrichFunction(
        FormulaAstNode node,
        FormulaAstNode root,
        IFormulaEvaluationContext context,
        string worksheetName,
        string sourceFormula)
    {
        if (node.IsStructural) return;

        var upper = node.Label.ToUpperInvariant();
        switch (upper)
        {
            case "IF":
            case "IFS":
            case "CHOOSE":
            case "SWITCH":
                if (node.Children.Count > 0)
                {
                    var branches = new List<string>();
                    foreach (var child in node.Children) branches.Add(child.Label);
                    var active = context.ResolveActiveIfBranch(node.Label, branches);
                    FormulaAstParser.SetActiveIfBranch(node, active);
                }

                break;
            case "INDEX":
                if (node.Children.Count >= 1 && node.Children[0].Kind == FormulaNodeKind.Reference)
                {
                    node.Children[0].Label = "array " + node.Children[0].Label;
                }

                break;
        }

        var subFormula = BuildSubFormula(node, sourceFormula);
        var evaluated = context.EvaluateSubExpression(subFormula, worksheetName);
        if (TraceUtils.IsUsableFormulaResult(evaluated))
        {
            node.Value = evaluated;
        }
        else if (IsEntireFormula(node, root))
        {
            var inherited = TraceUtils.FormatTraceValue(root.Value);
            if (TraceUtils.IsUsableFormulaResult(inherited))
            {
                node.Value = inherited;
            }
        }
        else if (!string.IsNullOrEmpty(evaluated))
        {
            node.Value = evaluated;
        }
    }

    private static bool IsEntireFormula(FormulaAstNode node, FormulaAstNode root) =>
        root.Children.Count == 1 && root.Children[0].Id == node.Id;

    private static string BuildSubFormula(FormulaAstNode node, string sourceFormula)
    {
        if (node.SourceLength > 0 &&
            node.SourceStart >= 0 &&
            node.SourceStart + node.SourceLength <= sourceFormula.Length)
        {
            var slice = sourceFormula.Substring(node.SourceStart, node.SourceLength);
            return slice.StartsWith("=", StringComparison.Ordinal) ? slice : "=" + slice;
        }

        return "=" + node.Label + "(" + string.Join(",", node.Children.ConvertAll(ChildExpression)) + ")";
    }

    private static string ChildExpression(FormulaAstNode child)
    {
        if (child.IsStructural)
        {
            return string.Join(",", child.Children.ConvertAll(ChildExpression));
        }

        if (child.Kind == FormulaNodeKind.Function && child.Children.Count > 0)
        {
            return child.Label + "(" + string.Join(",", child.Children.ConvertAll(ChildExpression)) + ")";
        }

        return child.Label;
    }
}
