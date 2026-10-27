using System;
using System.Collections.Generic;

namespace ArixcelExplorer.Core.Tracing;

public sealed class DependentsTreeNode
{
    public string Address { get; set; } = "";
    public string Display { get; set; } = "";
    public string Value { get; set; } = "";
    public string SheetName { get; set; } = "";
    public int Count { get; set; }
    public bool IsRoot { get; set; }
    public bool IsGroup { get; set; }
    public bool IsExpanded { get; set; } = true;
    public List<DependentsTreeNode> Children { get; } = new();
}

/// <summary>
/// Arixcel's Dependents window is a tree under the source cell. Each child is one
/// tracer-arrow target (G48:G53, or a single cell, or another sheet's range).
/// Adjacent cells are not merged just because they touch: G97, G98, and G99:G100
/// stay separate when Excel reports them as separate targets.
/// </summary>
public static class DependentsTree
{
    public static DependentsTreeNode Build(
        string? originAddress,
        string? originValue,
        IEnumerable<DependentEntry>? entries,
        bool groupBySheet)
    {
        var origin = NormalizeAddress(originAddress, originAddress);
        var leaves = new List<DependentsTreeNode>();
        if (entries != null)
        {
            foreach (var entry in entries)
            {
                var address = NormalizeAddress(entry?.Address, originAddress);
                if (string.IsNullOrEmpty(address)) continue;
                if (SameRange(address, origin)) continue;
                if (ContainsAddress(leaves, address)) continue;

                leaves.Add(new DependentsTreeNode
                {
                    Address = address,
                    Display = DisplayFor(address, originAddress, forceShort: false),
                    Value = entry?.Value ?? "",
                    SheetName = SheetOf(address),
                    IsExpanded = true
                });
            }
        }

        var root = new DependentsTreeNode
        {
            Address = string.IsNullOrEmpty(origin) ? originAddress ?? "" : origin,
            Display = DisplayFor(string.IsNullOrEmpty(origin) ? originAddress : origin, originAddress, forceShort: true),
            Value = originValue ?? "",
            SheetName = SheetOf(origin),
            IsRoot = true,
            IsExpanded = true,
            Count = leaves.Count
        };
        if (string.IsNullOrEmpty(root.Display))
        {
            root.Display = RangeOf(root.Address);
        }

        if (!groupBySheet)
        {
            root.Children.AddRange(leaves);
            return root;
        }

        var groups = new List<DependentsTreeNode>();
        foreach (var leaf in leaves)
        {
            DependentsTreeNode? group = null;
            foreach (var existing in groups)
            {
                if (SameSheet(existing.SheetName, leaf.SheetName))
                {
                    group = existing;
                    break;
                }
            }

            if (group == null)
            {
                group = new DependentsTreeNode
                {
                    SheetName = leaf.SheetName,
                    Display = string.IsNullOrEmpty(leaf.SheetName) ? "(sheet)" : leaf.SheetName,
                    IsGroup = true,
                    IsExpanded = true
                };
                groups.Add(group);
            }

            leaf.Display = DisplayFor(leaf.Address, originAddress, forceShort: true);
            group.Children.Add(leaf);
            group.Count = group.Children.Count;
        }

        root.Children.AddRange(groups);
        return root;
    }

    /// <summary>
    /// Merge tracer-arrow targets with DirectDependents areas.
    /// Arrow splits win: a geometric block that would swallow G97, G98, and G99:G100
    /// is dropped. Areas only fill ranges the arrows did not already cover.
    /// Other-sheet targets are kept after the same-sheet list.
    /// </summary>
    public static IReadOnlyList<DependentEntry> Combine(
        string? originAddress,
        IEnumerable<string>? arrowAddresses,
        IEnumerable<string>? areaAddresses)
    {
        var origin = NormalizeAddress(originAddress, originAddress);
        var originSheet = SheetOf(origin);
        var sameArrows = new List<string>();
        var other = new List<string>();
        Collect(arrowAddresses, originAddress, origin, originSheet, sameArrows, other);

        var sameAreas = new List<string>();
        if (areaAddresses != null)
        {
            foreach (var raw in areaAddresses)
            {
                var address = NormalizeAddress(raw, originAddress);
                if (string.IsNullOrEmpty(address) || SameRange(address, origin)) continue;
                if (string.IsNullOrEmpty(originSheet) || SameSheet(SheetOf(address), originSheet))
                {
                    sameAreas.Add(address);
                }
                else
                {
                    other.Add(address);
                }
            }
        }

        var same = new List<string>();
        if (sameArrows.Count > 0)
        {
            AddFiltered(same, sameArrows);
            AddFiltered(same, sameAreas);
        }
        else
        {
            AddFiltered(same, sameAreas);
        }

        var ordered = new List<string>(same);
        AddFiltered(ordered, other);

        var result = new List<DependentEntry>(ordered.Count);
        foreach (var address in ordered)
        {
            result.Add(new DependentEntry { Address = address, Count = 1 });
        }

        return result;
    }

    public static string NormalizeAddress(string? address, string? originAddress)
    {
        if (string.IsNullOrWhiteSpace(address)) return "";
        var trimmed = address!.Trim();
        var qualified = TraceUtils.QualifyAddress(trimmed, originAddress ?? "");
        if (string.IsNullOrEmpty(qualified)) qualified = trimmed;
        var parsed = TraceUtils.ParseWorksheetScopedAddress(qualified);
        if (parsed == null || !TraceUtils.IsRangeAddress(parsed.RangeAddress))
        {
            return TraceUtils.IsRangeAddress(trimmed) ? trimmed.Replace("$", "") : "";
        }

        var range = parsed.RangeAddress.Replace("$", "");
        if (string.IsNullOrEmpty(parsed.WorksheetName)) return range;
        return QuoteSheet(parsed.WorksheetName) + "!" + range;
    }

    public static bool ContainsRange(string? outer, string? inner)
    {
        if (!TryBounds(outer, out var outerSheet, out var outerMinRow, out var outerMaxRow, out var outerMinCol, out var outerMaxCol))
        {
            return false;
        }

        if (!TryBounds(inner, out var innerSheet, out var innerMinRow, out var innerMaxRow, out var innerMinCol, out var innerMaxCol))
        {
            return false;
        }

        if (!SameSheet(outerSheet, innerSheet)) return false;
        if (outerMinRow == innerMinRow && outerMaxRow == innerMaxRow &&
            outerMinCol == innerMinCol && outerMaxCol == innerMaxCol)
        {
            return false;
        }

        return innerMinRow >= outerMinRow && innerMaxRow <= outerMaxRow &&
               innerMinCol >= outerMinCol && innerMaxCol <= outerMaxCol;
    }

    private static void Collect(
        IEnumerable<string>? addresses,
        string? originAddress,
        string origin,
        string originSheet,
        List<string> sameSheet,
        List<string> otherSheet)
    {
        if (addresses == null) return;
        foreach (var raw in addresses)
        {
            var address = NormalizeAddress(raw, originAddress);
            if (string.IsNullOrEmpty(address) || SameRange(address, origin)) continue;
            if (string.IsNullOrEmpty(originSheet) || SameSheet(SheetOf(address), originSheet))
            {
                sameSheet.Add(address);
            }
            else
            {
                otherSheet.Add(address);
            }
        }
    }

    private static void AddFiltered(List<string> target, List<string> incoming)
    {
        foreach (var address in incoming)
        {
            var skip = false;
            foreach (var existing in target)
            {
                if (SameRange(existing, address) || ContainsRange(existing, address) || ContainsRange(address, existing))
                {
                    skip = true;
                    break;
                }
            }

            if (!skip) target.Add(address);
        }
    }

    private static bool ContainsAddress(List<DependentsTreeNode> leaves, string address)
    {
        foreach (var leaf in leaves)
        {
            if (SameRange(leaf.Address, address)) return true;
        }

        return false;
    }

    private static string DisplayFor(string? address, string? originAddress, bool forceShort)
    {
        if (string.IsNullOrWhiteSpace(address)) return "";
        if (!forceShort)
        {
            var formatted = TraceUtils.FormatExplorerLocation(address, originAddress ?? "");
            if (!string.IsNullOrEmpty(formatted)) return formatted;
        }

        return RangeOf(address!);
    }

    private static string RangeOf(string address)
    {
        var parsed = TraceUtils.ParseWorksheetScopedAddress(address);
        var range = parsed?.RangeAddress ?? address;
        return range.Replace("$", "");
    }

    private static string SheetOf(string address)
    {
        var parsed = TraceUtils.ParseWorksheetScopedAddress(address);
        return parsed?.WorksheetName ?? "";
    }

    private static bool SameRange(string left, string right) =>
        TraceUtils.AddressesReferToSameRange(left, right);

    private static bool SameSheet(string left, string right) =>
        string.Equals(left ?? "", right ?? "", StringComparison.OrdinalIgnoreCase);

    private static string QuoteSheet(string sheet) => "'" + sheet.Replace("'", "''") + "'";

    private static bool TryBounds(
        string? address,
        out string sheet,
        out int minRow,
        out int maxRow,
        out int minCol,
        out int maxCol)
    {
        sheet = "";
        minRow = maxRow = minCol = maxCol = 0;
        if (string.IsNullOrWhiteSpace(address)) return false;
        var parsed = TraceUtils.ParseWorksheetScopedAddress(address!);
        var range = (parsed?.RangeAddress ?? address!).Replace("$", "");
        if (!TraceUtils.IsRangeAddress(range)) return false;
        sheet = parsed?.WorksheetName ?? "";
        var parts = range.Split(':');
        var start = TraceUtils.ParseCellAddress(parts[0]);
        var end = TraceUtils.ParseCellAddress(parts.Length > 1 ? parts[1] : parts[0]);
        if (start == null || end == null) return false;
        minRow = Math.Min(start.Value.Row, end.Value.Row);
        maxRow = Math.Max(start.Value.Row, end.Value.Row);
        minCol = Math.Min(start.Value.Col, end.Value.Col);
        maxCol = Math.Max(start.Value.Col, end.Value.Col);
        return true;
    }
}
