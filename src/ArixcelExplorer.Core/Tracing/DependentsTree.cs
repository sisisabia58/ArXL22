using System;
using System.Collections.Generic;
using System.Globalization;

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
/// reported target: a row span such as D60:J60, a vertical fill such as G48:G53,
/// or a single cell. A reported row replaces the cells inside it. Separate targets
/// that only share a column, such as G97, G98, and G99:G100, stay separate.
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
        var prepared = new List<string>();
        if (entries != null)
        {
            foreach (var entry in entries)
            {
                var address = NormalizeAddress(entry?.Address, originAddress);
                if (string.IsNullOrEmpty(address) || SameRange(address, origin)) continue;
                prepared.Add(address);
            }
        }

        prepared = CoalesceAddresses(prepared);
        var leaves = new List<DependentsTreeNode>();
        foreach (var address in prepared)
        {
            leaves.Add(new DependentsTreeNode
            {
                Address = address,
                Display = DisplayFor(address, originAddress, forceShort: false),
                Value = ValueFor(entries, address),
                SheetName = SheetOf(address),
                IsExpanded = true
            });
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

            // Grouped rows stay sheet-qualified, matching Arixcel's
            // 'Summary (Instit)'!D60:J60 under the sheet node.
            leaf.Display = leaf.Address;
            group.Children.Add(leaf);
            group.Count = group.Children.Count;
        }

        root.Children.AddRange(groups);
        return root;
    }

    /// <summary>
    /// Merge tracer-arrow targets with DirectDependents areas.
    /// A reported row such as D60:J60 replaces single cells inside it.
    /// A block that would swallow distinct targets (G97:G100 over G99:G100) does not.
    /// Other-sheet targets stay after the same-sheet list.
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
        var coalesced = CoalesceAddresses(ordered);

        var result = new List<DependentEntry>(coalesced.Count);
        foreach (var address in coalesced)
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

    /// <summary>
    /// A single-row hit such as G60:J60 grows to the columns of a wider area on
    /// that same row (D60:J60, or row 60 of D60:J63). A vertical target such as
    /// G48:G53 is left unchanged.
    /// </summary>
    public static string WidenSingleRow(string? narrow, string? wide)
    {
        var left = NormalizeAddress(narrow, narrow);
        var right = NormalizeAddress(wide, string.IsNullOrEmpty(left) ? wide : left);
        if (string.IsNullOrEmpty(left)) return right;
        if (string.IsNullOrEmpty(right) || SameRange(left, right)) return left;
        if (!TryBounds(left, out var leftSheet, out var leftMinRow, out var leftMaxRow, out var leftMinCol, out var leftMaxCol))
        {
            return left;
        }

        if (!TryBounds(right, out var rightSheet, out var rightMinRow, out var rightMaxRow, out var rightMinCol, out var rightMaxCol))
        {
            return left;
        }

        if (leftMinRow != leftMaxRow || !SameSheet(leftSheet, rightSheet)) return left;
        if (rightMinRow > leftMinRow || rightMaxRow < leftMaxRow) return left;
        if (rightMinCol > leftMinCol || rightMaxCol < leftMaxCol) return left;
        return FormatA1(leftSheet, leftMinRow, rightMinCol, leftMaxRow, rightMaxCol);
    }

    public static bool FormulaReferencesOrigin(string? formula, string? contextSheet, string? originAddress)
    {
        if (string.IsNullOrWhiteSpace(formula)) return false;
        var text = formula!.Trim();
        if (text.IndexOf('=') < 0) return false;
        var origin = NormalizeAddress(originAddress, originAddress);
        if (string.IsNullOrEmpty(origin)) return false;
        var sheet = string.IsNullOrWhiteSpace(contextSheet) ? SheetOf(origin) : contextSheet!;
        foreach (var segment in Formulas.FormulaReferences.TokenizeFormulaReferences(text, sheet))
        {
            if (segment is not Formulas.FormulaReferenceSegment reference) continue;
            if (reference.Target is not TraceAddressTargetWrapper wrapper) continue;
            var target = NormalizeAddress(wrapper.Target.Address, origin);
            if (string.IsNullOrEmpty(target)) continue;
            if (SameRange(target, origin) || ContainsRange(target, origin)) return true;
        }

        return false;
    }

    public static IReadOnlyList<string> SplitAreas(string? address)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(address)) return result;
        var current = new System.Text.StringBuilder();
        var quoted = false;
        for (var i = 0; i < address!.Length; i++)
        {
            var ch = address[i];
            if (ch == '\'')
            {
                current.Append(ch);
                if (quoted && i + 1 < address.Length && address[i + 1] == '\'')
                {
                    current.Append(address[i + 1]);
                    i++;
                    continue;
                }

                quoted = !quoted;
                continue;
            }

            if (!quoted && (ch == ',' || ch == ';'))
            {
                var part = current.ToString().Trim();
                if (part.Length > 0) result.Add(part);
                current.Clear();
                continue;
            }

            current.Append(ch);
        }

        var last = current.ToString().Trim();
        if (last.Length > 0) result.Add(last);
        return result;
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
            // Keep a wider range even when smaller cells are already listed.
            // CoalesceAddresses decides which one survives.
            if (SameRange(existing, address) || ContainsRange(existing, address))
            {
                skip = true;
                break;
            }
            }

            if (!skip) target.Add(address);
        }
    }

    private static string ValueFor(IEnumerable<DependentEntry>? entries, string address)
    {
        if (entries == null) return "";
        foreach (var entry in entries)
        {
            if (entry == null) continue;
            if (SameRange(entry.Address, address)) return entry.Value ?? "";
        }

        return "";
    }

    private static List<string> CoalesceAddresses(List<string> incoming)
    {
        var items = new List<string>();
        foreach (var address in incoming)
        {
            if (string.IsNullOrEmpty(address)) continue;
            var duplicate = false;
            foreach (var existing in items)
            {
                if (SameRange(existing, address))
                {
                    duplicate = true;
                    break;
                }
            }

            if (!duplicate) items.Add(address);
        }

        RemoveWhere(items, address =>
        {
            if (IsSingleRow(address)) return false;
            foreach (var other in items)
            {
                if (SameRange(address, other) || IsSingleCell(other)) continue;
                if (ContainsRange(address, other)) return true;
            }

            return false;
        });

        RemoveWhere(items, address =>
        {
            if (!IsSingleCell(address)) return false;
            foreach (var other in items)
            {
                if (ContainsRange(other, address)) return true;
            }

            return false;
        });

        RemoveWhere(items, address =>
        {
            foreach (var other in items)
            {
                if (SameRange(address, other) || !IsSingleRow(other)) continue;
                if (ContainsRange(other, address)) return true;
            }

            return false;
        });

        RemoveWhere(items, address =>
        {
            if (IsSingleRow(address)) return false;
            foreach (var other in items)
            {
                if (!IsSingleRow(other) || IsSingleCell(other)) continue;
                if (Intersects(address, other)) return true;
            }

            return false;
        });

        return MergeTouchingRows(items);
    }

    private static void RemoveWhere(List<string> items, Func<string, bool> predicate)
    {
        for (var i = items.Count - 1; i >= 0; i--)
        {
            if (predicate(items[i])) items.RemoveAt(i);
        }
    }

    private static List<string> MergeTouchingRows(List<string> items)
    {
        var placed = new List<RowPiece>();
        var buckets = new Dictionary<string, List<RowPiece>>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < items.Count; i++)
        {
            if (!TryBounds(items[i], out var sheet, out var minRow, out var maxRow, out var minCol, out var maxCol) ||
                minRow != maxRow)
            {
                placed.Add(new RowPiece { Index = i, Address = items[i] });
                continue;
            }

            var key = sheet + "\n" + minRow.ToString(CultureInfo.InvariantCulture);
            if (!buckets.TryGetValue(key, out var bucket))
            {
                bucket = new List<RowPiece>();
                buckets[key] = bucket;
            }

            bucket.Add(new RowPiece
            {
                Index = i,
                Sheet = sheet,
                Row = minRow,
                MinCol = minCol,
                MaxCol = maxCol
            });
        }

        foreach (var bucket in buckets.Values)
        {
            bucket.Sort((left, right) => left.MinCol.CompareTo(right.MinCol));
            var current = bucket[0];
            for (var n = 1; n < bucket.Count; n++)
            {
                var next = bucket[n];
                if (next.MinCol <= current.MaxCol + 1)
                {
                    if (next.MaxCol > current.MaxCol) current.MaxCol = next.MaxCol;
                    if (next.Index < current.Index) current.Index = next.Index;
                    continue;
                }

                placed.Add(FinishRow(current));
                current = next;
            }

            placed.Add(FinishRow(current));
        }

        placed.Sort((left, right) => left.Index.CompareTo(right.Index));
        var result = new List<string>(placed.Count);
        foreach (var piece in placed)
        {
            if (!string.IsNullOrEmpty(piece.Address)) result.Add(piece.Address);
        }

        return result;
    }

    private static RowPiece FinishRow(RowPiece piece)
    {
        piece.Address = FormatA1(piece.Sheet, piece.Row, piece.MinCol, piece.Row, piece.MaxCol);
        return piece;
    }

    private static string FormatA1(string sheet, int row, int startColumn, int endRow, int endColumn)
    {
        var start = TraceUtils.ColumnLetters(startColumn) + (row + 1).ToString(CultureInfo.InvariantCulture);
        var end = TraceUtils.ColumnLetters(endColumn) + (endRow + 1).ToString(CultureInfo.InvariantCulture);
        var a1 = start == end ? start : start + ":" + end;
        if (string.IsNullOrEmpty(sheet)) return a1;
        return QuoteSheet(sheet) + "!" + a1;
    }

    private static bool IsSingleCell(string address)
    {
        if (!TryBounds(address, out _, out var minRow, out var maxRow, out var minCol, out var maxCol)) return false;
        return minRow == maxRow && minCol == maxCol;
    }

    private static bool IsSingleRow(string address)
    {
        if (!TryBounds(address, out _, out var minRow, out var maxRow, out _, out _)) return false;
        return minRow == maxRow;
    }

    private static bool Intersects(string left, string right)
    {
        if (!TryBounds(left, out var leftSheet, out var leftMinRow, out var leftMaxRow, out var leftMinCol, out var leftMaxCol))
        {
            return false;
        }

        if (!TryBounds(right, out var rightSheet, out var rightMinRow, out var rightMaxRow, out var rightMinCol, out var rightMaxCol))
        {
            return false;
        }

        if (!SameSheet(leftSheet, rightSheet)) return false;
        return leftMinRow <= rightMaxRow && rightMinRow <= leftMaxRow &&
               leftMinCol <= rightMaxCol && rightMinCol <= leftMaxCol;
    }

    private sealed class RowPiece
    {
        public int Index { get; set; }
        public string Address { get; set; } = "";
        public string Sheet { get; set; } = "";
        public int Row { get; set; }
        public int MinCol { get; set; }
        public int MaxCol { get; set; }
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
