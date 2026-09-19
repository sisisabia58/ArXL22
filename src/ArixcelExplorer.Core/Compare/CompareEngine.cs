using System;
using System.Collections.Generic;

namespace ArixcelExplorer.Core.Compare;

public sealed class CompareCellSnapshot
{
    public string Address { get; set; } = "";
    public object? Value { get; set; }
    public string? Formula { get; set; }
    public string? NumberFormat { get; set; }
}

public sealed class CompareDiffEntry
{
    public string LeftAddress { get; set; } = "";
    public string RightAddress { get; set; } = "";
    public string LeftValue { get; set; } = "";
    public string RightValue { get; set; } = "";
    public string DiffKind { get; set; } = "";
}

public sealed class CompareAlignmentResult
{
    public IReadOnlyList<int> LeftRowMapping { get; set; } = Array.Empty<int>();
    public IReadOnlyList<int> RightRowMapping { get; set; } = Array.Empty<int>();
}

public static class CompareEngine
{
    public static CompareAlignmentResult AlignByContent(
        IReadOnlyList<string> leftRowSignatures,
        IReadOnlyList<string> rightRowSignatures)
    {
        var leftMap = new List<int>();
        var rightUsed = new HashSet<int>();
        for (var li = 0; li < leftRowSignatures.Count; li++)
        {
            var sig = leftRowSignatures[li];
            var match = -1;
            for (var ri = 0; ri < rightRowSignatures.Count; ri++)
            {
                if (rightUsed.Contains(ri)) continue;
                if (rightRowSignatures[ri] == sig)
                {
                    match = ri;
                    rightUsed.Add(ri);
                    break;
                }
            }

            leftMap.Add(match);
        }

        var rightMap = new List<int>();
        for (var ri = 0; ri < rightRowSignatures.Count; ri++)
        {
            var found = -1;
            for (var li = 0; li < leftMap.Count; li++)
            {
                if (leftMap[li] == ri)
                {
                    found = li;
                    break;
                }
            }

            rightMap.Add(found);
        }

        return new CompareAlignmentResult
        {
            LeftRowMapping = leftMap,
            RightRowMapping = rightMap
        };
    }

    public static IReadOnlyList<CompareDiffEntry> DiffGrids(
        CompareCellSnapshot[][] left,
        CompareCellSnapshot[][] right,
        string leftPrefix,
        string rightPrefix)
    {
        var diffs = new List<CompareDiffEntry>();
        var maxRows = Math.Max(left.Length, right.Length);
        var maxCols = 0;
        foreach (var row in left) maxCols = Math.Max(maxCols, row.Length);
        foreach (var row in right) maxCols = Math.Max(maxCols, row.Length);

        for (var r = 0; r < maxRows; r++)
        {
            for (var c = 0; c < maxCols; c++)
            {
                var l = r < left.Length && c < left[r].Length ? left[r][c] : null;
                var rt = r < right.Length && c < right[r].Length ? right[r][c] : null;
                var leftVal = FormatValue(l?.Value);
                var rightVal = FormatValue(rt?.Value);
                if (leftVal == rightVal && (l?.Formula ?? "") == (rt?.Formula ?? "")) continue;

                diffs.Add(new CompareDiffEntry
                {
                    LeftAddress = l?.Address ?? $"{leftPrefix}R{r + 1}C{c + 1}",
                    RightAddress = rt?.Address ?? $"{rightPrefix}R{r + 1}C{c + 1}",
                    LeftValue = leftVal,
                    RightValue = rightVal,
                    DiffKind = l == null ? "MissingLeft" : rt == null ? "MissingRight" : "ValueOrFormula"
                });
            }
        }

        return diffs;
    }

    private static string FormatValue(object? value)
    {
        if (value == null) return "";
        return value.ToString() ?? "";
    }
}
