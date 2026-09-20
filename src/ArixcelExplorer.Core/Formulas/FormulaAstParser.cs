using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ArixcelExplorer.Core.Tracing;

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
    public bool FlattenIntoParent { get; set; }
    public string? ParentId { get; set; }
    public int SourceStart { get; set; }
    public int SourceLength { get; set; }
    public List<FormulaAstNode> Children { get; set; } = new();
}

public static class FormulaAstParser
{
    private static readonly Regex RangeToken = new(
        @"^(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_\\][A-Za-z0-9_.]*)!)?(?<range>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex NumberToken = new(
        @"^(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?",
        RegexOptions.Compiled);

    private static readonly Regex NameToken = new(
        @"^[A-Za-z_\\][A-Za-z0-9_.]*",
        RegexOptions.Compiled);

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
            var (node, _) = ParseComparison(normalized, 0, root.Id, cellAddress);
            AttachTopLevel(root, node);
            QualifyReferences(root, cellAddress);
        }
        catch
        {
            root.Children.Add(new FormulaAstNode
            {
                Kind = FormulaNodeKind.Error,
                Label = normalized,
                ParentId = root.Id,
                SourceStart = 0,
                SourceLength = normalized.Length
            });
        }

        ApplyDisplayOffset(root, DisplayPrefixLength(formula));
        return root;
    }

    public static void AttachValidationSource(FormulaAstNode root, ValidationListSource source)
    {
        var kind = source.IsNamedRange ? FormulaNodeKind.NamedRange : FormulaNodeKind.Reference;
        root.Children.Add(new FormulaAstNode
        {
            Kind = kind,
            Label = source.Label,
            Info = "validation",
            Location = source.Location,
            ParentId = root.Id
        });
    }

    public static bool LooksLikeReference(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text.Trim();
        var match = RangeToken.Match(trimmed);
        return match.Success && match.Length == trimmed.Length && TraceUtils.IsRangeAddress(match.Groups["range"].Value);
    }

    private static int DisplayPrefixLength(string formula)
    {
        var trimmed = formula.Trim();
        if (trimmed.StartsWith("{=", StringComparison.Ordinal)) return 2;
        if (trimmed.StartsWith("=", StringComparison.Ordinal)) return 1;
        return 0;
    }

    private static void ApplyDisplayOffset(FormulaAstNode node, int prefix)
    {
        if (prefix != 0 && node.Kind != FormulaNodeKind.Root && node.SourceLength > 0)
        {
            node.SourceStart += prefix;
        }

        foreach (var child in node.Children)
        {
            ApplyDisplayOffset(child, prefix);
        }
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

    private static void AttachTopLevel(FormulaAstNode root, FormulaAstNode node)
    {
        if (node.FlattenIntoParent && node.Children.Count > 0)
        {
            foreach (var child in node.Children)
            {
                child.ParentId = root.Id;
                root.Children.Add(child);
            }

            return;
        }

        node.ParentId = root.Id;
        root.Children.Add(node);
    }

    private static void QualifyReferences(FormulaAstNode node, string cellAddress)
    {
        if (node.Kind is FormulaNodeKind.Reference or FormulaNodeKind.NamedRange)
        {
            var raw = string.IsNullOrWhiteSpace(node.Location) ? node.Label : node.Location ?? "";
            var qualified = TraceUtils.QualifyAddress(raw, cellAddress);
            if (!string.IsNullOrEmpty(qualified))
            {
                node.Location = qualified;
            }
        }

        foreach (var child in node.Children)
        {
            QualifyReferences(child, cellAddress);
        }
    }

    private static (FormulaAstNode Node, int NextIndex) ParseComparison(string input, int start, string? parentId, string contextLocation)
    {
        var (left, index) = ParseConcat(input, start, parentId, contextLocation);
        index = SkipSpaces(input, index);
        if (!TryReadComparison(input, index, out var op, out var afterOp))
        {
            return (left, index);
        }

        var (right, next) = ParseConcat(input, afterOp, parentId, contextLocation);
        var sourceStart = left.SourceStart;
        var sourceEnd = SourceEnd(right);
        var node = new FormulaAstNode
        {
            Kind = FormulaNodeKind.Operator,
            Label = input.Substring(sourceStart, sourceEnd - sourceStart),
            ParentId = parentId,
            SourceStart = sourceStart,
            SourceLength = sourceEnd - sourceStart,
            Children = { left, right }
        };
        left.ParentId = node.Id;
        right.ParentId = node.Id;
        _ = op;
        return (node, next);
    }

    private static (FormulaAstNode Node, int NextIndex) ParseConcat(string input, int start, string? parentId, string contextLocation)
    {
        return ParseBinary(input, start, parentId, contextLocation, ParseAdd, flatten: false, "&");
    }

    private static (FormulaAstNode Node, int NextIndex) ParseAdd(string input, int start, string? parentId, string contextLocation)
    {
        return ParseBinary(input, start, parentId, contextLocation, ParseMul, flatten: true, "+", "-");
    }

    private static (FormulaAstNode Node, int NextIndex) ParseMul(string input, int start, string? parentId, string contextLocation)
    {
        return ParseBinary(input, start, parentId, contextLocation, ParsePower, flatten: true, "*", "/");
    }

    private static (FormulaAstNode Node, int NextIndex) ParsePower(string input, int start, string? parentId, string contextLocation)
    {
        return ParseBinary(input, start, parentId, contextLocation, ParsePercent, flatten: false, "^");
    }

    private static (FormulaAstNode Node, int NextIndex) ParseBinary(
        string input,
        int start,
        string? parentId,
        string contextLocation,
        Func<string, int, string?, string, (FormulaAstNode Node, int NextIndex)> parseOperand,
        bool flatten,
        params string[] operators)
    {
        var (first, index) = parseOperand(input, start, parentId, contextLocation);
        var factors = new List<FormulaAstNode> { first };
        while (true)
        {
            index = SkipSpaces(input, index);
            if (!TryReadOperator(input, index, operators, out _, out var afterOp))
            {
                break;
            }

            var (nextNode, nextIndex) = parseOperand(input, afterOp, parentId, contextLocation);
            factors.Add(nextNode);
            index = nextIndex;
        }

        if (factors.Count == 1)
        {
            return (first, index);
        }

        var sourceStart = factors[0].SourceStart;
        var sourceEnd = SourceEnd(factors[factors.Count - 1]);
        var group = new FormulaAstNode
        {
            Kind = FormulaNodeKind.Operator,
            Label = input.Substring(sourceStart, sourceEnd - sourceStart),
            ParentId = parentId,
            FlattenIntoParent = flatten,
            SourceStart = sourceStart,
            SourceLength = sourceEnd - sourceStart,
            Children = factors
        };
        foreach (var child in factors)
        {
            child.ParentId = group.Id;
        }

        return (group, index);
    }

    private static (FormulaAstNode Node, int NextIndex) ParsePercent(string input, int start, string? parentId, string contextLocation)
    {
        var (node, index) = ParseUnary(input, start, parentId, contextLocation);
        index = SkipSpaces(input, index);
        if (index < input.Length && input[index] == '%')
        {
            node.Label += "%";
            node.SourceLength = index + 1 - node.SourceStart;
            index += 1;
        }

        return (node, index);
    }

    private static (FormulaAstNode Node, int NextIndex) ParseUnary(string input, int start, string? parentId, string contextLocation)
    {
        var index = SkipSpaces(input, start);
        if (index < input.Length && (input[index] == '+' || input[index] == '-'))
        {
            var sign = input[index];
            var (inner, next) = ParseUnary(input, index + 1, parentId, contextLocation);
            if (sign == '+')
            {
                return (inner, next);
            }

            var node = new FormulaAstNode
            {
                Kind = FormulaNodeKind.Operator,
                Label = input.Substring(index, SourceEnd(inner) - index),
                ParentId = parentId,
                SourceStart = index,
                SourceLength = SourceEnd(inner) - index,
                Children = { inner }
            };
            inner.ParentId = node.Id;
            return (node, next);
        }

        return ParsePrimary(input, index, parentId, contextLocation);
    }

    private static (FormulaAstNode Node, int NextIndex) ParsePrimary(string input, int start, string? parentId, string contextLocation)
    {
        var index = SkipSpaces(input, start);
        if (TryParseFunction(input, index, parentId, contextLocation, out var functionNode, out var functionEnd))
        {
            return (functionNode, functionEnd);
        }

        if (index < input.Length && input[index] == '(')
        {
            var (inner, afterInner) = ParseComparison(input, index + 1, parentId, contextLocation);
            var close = SkipSpaces(input, afterInner);
            if (close < input.Length && input[close] == ')')
            {
                close += 1;
            }

            inner.ParentId = parentId;
            return (inner, close);
        }

        if (TryParseReference(input, index, parentId, contextLocation, out var reference, out var refEnd))
        {
            return (reference, refEnd);
        }

        if (TryParseNumber(input, index, parentId, out var number, out var numEnd))
        {
            return (number, numEnd);
        }

        if (TryParseString(input, index, parentId, out var text, out var textEnd))
        {
            return (text, textEnd);
        }

        if (TryParseName(input, index, parentId, contextLocation, out var named, out var nameEnd))
        {
            return (named, nameEnd);
        }

        var leftover = input.Substring(index);
        return (new FormulaAstNode
        {
            Kind = FormulaNodeKind.Literal,
            Label = leftover,
            Value = leftover,
            ParentId = parentId,
            SourceStart = index,
            SourceLength = leftover.Length
        }, input.Length);
    }

    private static bool TryParseFunction(
        string input,
        int start,
        string? parentId,
        string contextLocation,
        out FormulaAstNode node,
        out int end)
    {
        node = new FormulaAstNode { ParentId = parentId };
        end = start;
        if (start >= input.Length) return false;
        var nameMatch = Regex.Match(input.Substring(start), @"^([A-Za-z_][A-Za-z0-9_.]*)\(");
        if (!nameMatch.Success)
        {
            return false;
        }

        var name = nameMatch.Groups[1].Value.ToUpperInvariant();
        var openParen = start + nameMatch.Length - 1;
        node.Kind = FormulaNodeKind.Function;
        node.Label = name;
        node.SourceStart = start;
        var args = ParseArgumentList(input, openParen + 1, node.Id, contextLocation);
        node.Children = args.Nodes;
        for (var i = 0; i < node.Children.Count; i++)
        {
            node.Children[i].Info = FunctionArgInfo.LabelFor(name, i);
        }

        end = args.EndIndex + 1;
        node.SourceLength = end - start;
        return true;
    }

    private static (List<FormulaAstNode> Nodes, int EndIndex) ParseArgumentList(
        string input,
        int start,
        string functionId,
        string contextLocation)
    {
        var nodes = new List<FormulaAstNode>();
        var index = start;
        var depth = 1;
        var argStart = start;
        var inString = false;

        while (index < input.Length && depth > 0)
        {
            var ch = input[index];
            if (inString)
            {
                if (ch == '"')
                {
                    if (index + 1 < input.Length && input[index + 1] == '"')
                    {
                        index += 2;
                        continue;
                    }

                    inString = false;
                }

                index += 1;
                continue;
            }

            if (ch == '"')
            {
                inString = true;
                index += 1;
                continue;
            }

            if (ch == '(') depth += 1;
            else if (ch == ')')
            {
                depth -= 1;
                if (depth == 0)
                {
                    var slice = ParseSlice(input, argStart, index, functionId, contextLocation);
                    if (slice != null) nodes.Add(slice);
                    return (nodes, index);
                }
            }
            else if (ch == ',' && depth == 1)
            {
                var slice = ParseSlice(input, argStart, index, functionId, contextLocation);
                if (slice != null) nodes.Add(slice);
                argStart = index + 1;
            }

            index += 1;
        }

        return (nodes, index);
    }

    private static FormulaAstNode? ParseSlice(
        string input,
        int start,
        int endExclusive,
        string? parentId,
        string contextLocation)
    {
        var sliceStart = SkipSpaces(input, start);
        var sliceEnd = endExclusive;
        while (sliceEnd > sliceStart && char.IsWhiteSpace(input[sliceEnd - 1])) sliceEnd -= 1;
        if (sliceEnd <= sliceStart) return null;

        var text = input.Substring(sliceStart, sliceEnd - sliceStart);
        var (node, _) = ParseComparison(text, 0, parentId, contextLocation);
        OffsetSources(node, sliceStart);
        return node;
    }

    private static void OffsetSources(FormulaAstNode node, int offset)
    {
        node.SourceStart += offset;
        foreach (var child in node.Children)
        {
            OffsetSources(child, offset);
        }
    }

    private static bool TryParseReference(
        string input,
        int start,
        string? parentId,
        string contextLocation,
        out FormulaAstNode node,
        out int end)
    {
        node = new FormulaAstNode();
        end = start;
        if (start >= input.Length) return false;
        var match = RangeToken.Match(input.Substring(start));
        if (!match.Success || !TraceUtils.IsRangeAddress(match.Groups["range"].Value))
        {
            return false;
        }

        var nextChar = start + match.Length < input.Length ? input[start + match.Length] : '\0';
        if (nextChar == '(') return false;

        var text = match.Value;
        node = new FormulaAstNode
        {
            Kind = FormulaNodeKind.Reference,
            Label = text,
            Location = TraceUtils.QualifyAddress(text, contextLocation),
            ParentId = parentId,
            SourceStart = start,
            SourceLength = match.Length
        };
        end = start + match.Length;
        return true;
    }

    private static bool TryParseNumber(string input, int start, string? parentId, out FormulaAstNode node, out int end)
    {
        node = new FormulaAstNode();
        end = start;
        if (start >= input.Length) return false;
        var match = NumberToken.Match(input.Substring(start));
        if (!match.Success) return false;

        node = new FormulaAstNode
        {
            Kind = FormulaNodeKind.Literal,
            Label = match.Value,
            Value = match.Value,
            ParentId = parentId,
            SourceStart = start,
            SourceLength = match.Length
        };
        end = start + match.Length;
        return true;
    }

    private static bool TryParseString(string input, int start, string? parentId, out FormulaAstNode node, out int end)
    {
        node = new FormulaAstNode();
        end = start;
        if (start >= input.Length || input[start] != '"') return false;
        var index = start + 1;
        while (index < input.Length)
        {
            if (input[index] == '"')
            {
                if (index + 1 < input.Length && input[index + 1] == '"')
                {
                    index += 2;
                    continue;
                }

                index += 1;
                break;
            }

            index += 1;
        }

        var text = input.Substring(start, index - start);
        node = new FormulaAstNode
        {
            Kind = FormulaNodeKind.Literal,
            Label = text,
            Value = text,
            ParentId = parentId,
            SourceStart = start,
            SourceLength = text.Length
        };
        end = index;
        return true;
    }

    private static bool TryParseName(
        string input,
        int start,
        string? parentId,
        string contextLocation,
        out FormulaAstNode node,
        out int end)
    {
        node = new FormulaAstNode();
        end = start;
        if (start >= input.Length) return false;
        var match = NameToken.Match(input.Substring(start));
        if (!match.Success) return false;
        var next = start + match.Length < input.Length ? input[start + match.Length] : '\0';
        if (next == '(') return false;

        var text = match.Value;
        if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
            text.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            node = new FormulaAstNode
            {
                Kind = FormulaNodeKind.Literal,
                Label = text.ToUpperInvariant(),
                Value = text.ToUpperInvariant(),
                ParentId = parentId,
                SourceStart = start,
                SourceLength = match.Length
            };
            end = start + match.Length;
            return true;
        }

        node = new FormulaAstNode
        {
            Kind = FormulaNodeKind.NamedRange,
            Label = text,
            Location = TraceUtils.QualifyAddress(text, contextLocation),
            ParentId = parentId,
            SourceStart = start,
            SourceLength = match.Length
        };
        end = start + match.Length;
        return true;
    }

    private static bool TryReadComparison(string input, int index, out string op, out int next)
    {
        foreach (var candidate in new[] { ">=", "<=", "<>", ">", "<", "=" })
        {
            if (index + candidate.Length <= input.Length &&
                string.Compare(input, index, candidate, 0, candidate.Length, StringComparison.Ordinal) == 0)
            {
                op = candidate;
                next = SkipSpaces(input, index + candidate.Length);
                return true;
            }
        }

        op = "";
        next = index;
        return false;
    }

    private static bool TryReadOperator(string input, int index, string[] operators, out string op, out int next)
    {
        foreach (var candidate in operators)
        {
            if (index + candidate.Length <= input.Length &&
                string.Compare(input, index, candidate, 0, candidate.Length, StringComparison.Ordinal) == 0)
            {
                op = candidate;
                next = SkipSpaces(input, index + candidate.Length);
                return true;
            }
        }

        op = "";
        next = index;
        return false;
    }

    private static int SkipSpaces(string input, int index)
    {
        while (index < input.Length && char.IsWhiteSpace(input[index])) index += 1;
        return index;
    }

    private static int SourceEnd(FormulaAstNode node) => node.SourceStart + node.SourceLength;

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

    public static void CollapseFunctions(FormulaAstNode node)
    {
        if (node.Kind == FormulaNodeKind.Function)
        {
            node.IsExpanded = false;
        }

        foreach (var child in node.Children)
        {
            CollapseFunctions(child);
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
