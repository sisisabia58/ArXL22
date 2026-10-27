using System.Collections.Generic;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.Core.Formulas;

public static class ExplorerLocations
{
    public static string NavigationAddress(FormulaAstNode node, string rootAddress)
    {
        if (node.Kind == FormulaNodeKind.Operator)
        {
            var addresses = new List<string>();
            CollectReferences(node, addresses);
            return TraceUtils.BoundingAddress(addresses);
        }

        if (node.Kind is FormulaNodeKind.Literal or FormulaNodeKind.Error)
        {
            return "";
        }

        if (string.IsNullOrWhiteSpace(node.Location)) return "";
        return TraceUtils.QualifyAddress(node.Location!, rootAddress);
    }

    public static string DisplayValue(FormulaAstNode node, FormulaAstNode? root)
    {
        var formatted = TraceUtils.FormatTraceValue(node.Value);
        if (!string.IsNullOrEmpty(formatted)) return formatted;
        if (node.Kind == FormulaNodeKind.Operator && root != null && node.ParentId == root.Id)
        {
            return TraceUtils.FormatTraceValue(root.Value);
        }

        return "";
    }

    /// <summary>
    /// Arixcel opens =D4+E4 with the expression row selected, not the origin cell.
    /// </summary>
    public static int OpeningSelectionIndex(int visibleRowCount) => visibleRowCount > 1 ? 1 : 0;

    public static string DependentsStatus(int dependentCount)
    {
        if (dependentCount <= 0) return "No direct dependents found";
        if (dependentCount == 1) return "1 direct dependent";
        return dependentCount + " direct dependents";
    }

    private static void CollectReferences(FormulaAstNode node, List<string> addresses)
    {
        if (node.Kind is FormulaNodeKind.Reference or FormulaNodeKind.NamedRange &&
            !string.IsNullOrWhiteSpace(node.Location))
        {
            addresses.Add(node.Location!);
        }

        foreach (var child in node.Children)
        {
            CollectReferences(child, addresses);
        }
    }
}
