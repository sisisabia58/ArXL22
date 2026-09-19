using System;
using System.Collections.Generic;
using System.Text;

namespace ArixcelExplorer.Core.Formulas;

public enum FormulaNodeKind
{
    Root,
    Function,
    Operator,
    Reference,
    Literal,
    Array,
    NamedRange,
    Error
}

public sealed class FormulaAstNode
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public FormulaNodeKind Kind { get; set; }
    public string Label { get; set; } = "";
    public string Info { get; set; } = "";
    public string? Value { get; set; }
    public string? Location { get; set; }
    public bool IsActiveBranch { get; set; }
    public bool IsExpanded { get; set; } = true;
    public string? ParentId { get; set; }
    public List<FormulaAstNode> Children { get; set; } = new();
}

public static class FormulaAstParser
{
    public static FormulaAstNode Parse(string formula, string cellAddress, string? evaluatedValue = null)
    {
        var normalized = NormalizeFormula(formula);
        var root = new FormulaAstNode
        {
            Kind = FormulaNodeKind.Root,
            Label = cellAddress,
            Value = evaluatedValue,
            Location = cellAddress
        };

        if (string.IsNullOrWhiteSpace(normalized))
        {
            root.Kind = FormulaNodeKind.Literal;
            root.Label = evaluatedValue ?? "(constant)";
            return root;
        }

        try
        {
            var (node, _) = ParseExpression(normalized, root.Id, cellAddress);
            root.Children.Add(node);
            root.Label = cellAddress;
        }
        catch
        {
            root.Children.Add(new FormulaAstNode
            {
                Kind = FormulaNodeKind.Error,
                Label = normalized,
                ParentId = root.Id
            });
        }

        return root;
    }

    public static IReadOnlyList<FormulaAstNode> FlattenVisible(FormulaAstNode root)
    {
        var rows = new List<FormulaAstNode>();
        Flatten(root, 0, rows);
        return rows;
    }

    private static void Flatten(FormulaAstNode node, int depth, IList<FormulaAstNode> rows)
    {
        rows.Add(node);
        if (!node.IsExpanded) return;
        foreach (var child in node.Children)
        {
            Flatten(child, depth + 1, rows);
        }
    }

    private static string NormalizeFormula(string formula)
    {
        var trimmed = formula.Trim();
        if (trimmed.StartsWith("{=", StringComparison.Ordinal) && trimmed.EndsWith("}", StringComparison.Ordinal))
        {
            return trimmed.Substring(2, trimmed.Length - 3);
        }

        if (trimmed.StartsWith("=", StringComparison.Ordinal))
        {
            return trimmed.Substring(1);
        }

        return trimmed;
    }

    private static (FormulaAstNode Node, int NextIndex) ParseExpression(string input, string? parentId, string contextLocation)
    {
        SkipWhitespace(input, ref contextLocation);
        if (TryParseFunction(input, 0, parentId, out var functionNode, out var functionEnd))
        {
            return (functionNode, functionEnd);
        }

        return ParseAtom(input, 0, parentId, contextLocation);
    }

    private static bool TryParseFunction(string input, int start, string? parentId, out FormulaAstNode node, out int end)
    {
        node = new FormulaAstNode { ParentId = parentId };
        end = start;
        var nameMatch = System.Text.RegularExpressions.Regex.Match(input.Substring(start), @"^([A-Za-z_][A-Za-z0-9_.]*)\(");
        if (!nameMatch.Success)
        {
            return false;
        }

        var name = nameMatch.Groups[1].Value.ToUpperInvariant();
        var openParen = start + nameMatch.Length - 1;
        var args = ParseArgumentList(input, openParen + 1, parentId);
        node.Kind = FormulaNodeKind.Function;
        node.Label = name;
        node.Children = args.Nodes;
        for (var i = 0; i < node.Children.Count; i++)
        {
            node.Children[i].Info = FunctionArgInfo.LabelFor(name, i);
        }
        end = args.EndIndex + 1;
        return true;
    }

    private static (List<FormulaAstNode> Nodes, int EndIndex) ParseArgumentList(string input, int start, string? parentId)
    {
        var nodes = new List<FormulaAstNode>();
        var index = start;
        var depth = 1;
        var argStart = start;

        while (index < input.Length && depth > 0)
        {
            var ch = input[index];
            if (ch == '(') depth += 1;
            else if (ch == ')')
            {
                depth -= 1;
                if (depth == 0)
                {
                    var argText = input.Substring(argStart, index - argStart).Trim();
                    if (argText.Length > 0)
                    {
                        nodes.Add(ParseArgument(argText, parentId));
                    }

                    return (nodes, index);
                }
            }
            else if (ch == ',' && depth == 1)
            {
                var argText = input.Substring(argStart, index - argStart).Trim();
                if (argText.Length > 0)
                {
                    nodes.Add(ParseArgument(argText, parentId));
                }

                argStart = index + 1;
            }

            index += 1;
        }

        return (nodes, index);
    }

    private static FormulaAstNode ParseArgument(string text, string? parentId)
    {
        if (TryParseFunction(text, 0, parentId, out var fn, out _))
        {
            return fn;
        }

        if (LooksLikeReference(text))
        {
            return new FormulaAstNode
            {
                Kind = FormulaNodeKind.Reference,
                Label = text,
                Location = text,
                ParentId = parentId
            };
        }

        return new FormulaAstNode
        {
            Kind = FormulaNodeKind.Literal,
            Label = text,
            Value = text,
            ParentId = parentId
        };
    }

    private static (FormulaAstNode Node, int NextIndex) ParseAtom(string input, int start, string? parentId, string contextLocation)
    {
        var remaining = input.Substring(start).Trim();
        if (LooksLikeReference(remaining))
        {
            return (new FormulaAstNode
            {
                Kind = FormulaNodeKind.Reference,
                Label = remaining,
                Location = remaining,
                ParentId = parentId
            }, input.Length);
        }

        return (new FormulaAstNode
        {
            Kind = FormulaNodeKind.Literal,
            Label = remaining,
            Value = remaining,
            ParentId = parentId
        }, input.Length);
    }

    private static bool LooksLikeReference(string text)
    {
        return System.Text.RegularExpressions.Regex.IsMatch(
            text.Trim(),
            @"^(\[[^\]]+\])?('?[^'!]+'?!)?\$?[A-Za-z]{1,3}\$?\d",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    private static void SkipWhitespace(string input, ref string contextLocation)
    {
        _ = input;
        _ = contextLocation;
    }

    public static void SetActiveIfBranch(FormulaAstNode root, int chosenIndex)
    {
        if (root.Kind != FormulaNodeKind.Function) return;
        var upper = root.Label.ToUpperInvariant();
        if (upper is not ("IF" or "IFS" or "CHOOSE" or "SWITCH")) return;

        for (var i = 0; i < root.Children.Count; i++)
        {
            root.Children[i].IsActiveBranch = i == chosenIndex;
        }
    }

    public static void ToggleExpand(FormulaAstNode node) => node.IsExpanded = !node.IsExpanded;

    public static void ExpandAll(FormulaAstNode root)
    {
        root.IsExpanded = true;
        foreach (var child in root.Children)
        {
            ExpandAll(child);
        }
    }

    public static void CollapseAll(FormulaAstNode root)
    {
        if (root.Kind != FormulaNodeKind.Root)
        {
            root.IsExpanded = false;
        }

        foreach (var child in root.Children)
        {
            CollapseAll(child);
        }
    }

    /// <summary>
    /// Cycles ExpandAll → CollapseAll. Returns whether the tree is fully expanded after the call.
    /// CollapseAll keeps the origin/root node expanded so its direct children stay visible.
    /// </summary>
    public static bool CycleExpandAll(FormulaAstNode root, bool currentlyExpanded)
    {
        if (currentlyExpanded)
        {
            CollapseAll(root);
            return false;
        }

        ExpandAll(root);
        return true;
    }
}
