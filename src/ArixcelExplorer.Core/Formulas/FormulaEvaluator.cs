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
        EnrichNode(root, context, worksheetName);
    }

    private static void EnrichNode(FormulaAstNode node, IFormulaEvaluationContext context, string worksheetName)
    {
        foreach (var child in node.Children)
        {
            EnrichNode(child, context, worksheetName);
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

                if (node.Location != null && context.TryResolveReferenceValue(node.Location, out var refValue))
                {
                    node.Value = refValue;
                }

                break;
            case FormulaNodeKind.Function:
                EnrichFunction(node, context, worksheetName);
                break;
            case FormulaNodeKind.Operator:
                if (!string.IsNullOrWhiteSpace(node.Label))
                {
                    var evaluated = context.EvaluateSubExpression("=" + node.Label, worksheetName);
                    if (!string.IsNullOrEmpty(evaluated))
                    {
                        node.Value = evaluated;
                    }
                }

                break;
        }
    }

    private static void EnrichFunction(FormulaAstNode node, IFormulaEvaluationContext context, string worksheetName)
    {
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
            case "SUM":
            case "SUMIF":
            case "SUMIFS":
                node.Label = upper + $" ({node.Children.Count} args)";
                break;
        }

        var subFormula = "=" + node.Label + "(" + string.Join(", ", node.Children.ConvertAll(c => c.Label)) + ")";
        var evaluated = context.EvaluateSubExpression(subFormula, worksheetName);
        if (!string.IsNullOrEmpty(evaluated))
        {
            node.Value = evaluated;
        }
    }
}
