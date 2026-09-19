using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ArixcelExplorer.Core.Tracing;

public sealed class TraceRow
{
    public int Level { get; set; }
    public string Address { get; set; } = "";
    public string Value { get; set; } = "";
    public string Formula { get; set; } = "";
    public string? ParentAddress { get; set; }
}

public sealed class TraceCellInfo
{
    public string WorksheetName { get; set; } = "";
    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
    public string Address { get; set; } = "";
    public object? Value { get; set; }
    public object? Formula { get; set; }
}

public sealed class TraceBuilderProgress
{
    public IReadOnlyList<TraceRow> Rows { get; set; } = Array.Empty<TraceRow>();
    public int Level { get; set; }
    public bool IsFinal { get; set; }
    public bool Truncated { get; set; }
}

public sealed class TraceBuilderResult
{
    public IReadOnlyList<TraceRow> Rows { get; set; } = Array.Empty<TraceRow>();
    public bool Truncated { get; set; }
}

public sealed class TraceBuilderInput
{
    public TraceCellInfo Root { get; set; } = new();
    public int MaxDepth { get; set; }
    public int? MaxRows { get; set; }
    public Func<IReadOnlyList<TraceCellInfo>, Task<IReadOnlyList<IReadOnlyList<TraceCellInfo>>>> GetAllNeighbors { get; set; } =
        _ => Task.FromResult<IReadOnlyList<IReadOnlyList<TraceCellInfo>>>(Array.Empty<IReadOnlyList<TraceCellInfo>>());
    public Func<TraceBuilderProgress, Task>? OnProgress { get; set; }
}

public static class TraceBuilder
{
    public static TraceRow ToTraceRow(TraceCellInfo cell, int level, string? parentAddress)
    {
        return new TraceRow
        {
            Level = level,
            Address = cell.Address,
            Value = TraceUtils.FormatTraceValue(cell.Value),
            Formula = TraceUtils.FormatTraceFormula(cell.Formula),
            ParentAddress = parentAddress
        };
    }

    public static async Task<TraceBuilderResult> BuildAsync(TraceBuilderInput input)
    {
        var maxRows = input.MaxRows ?? TraceUtils.MaxTraceRows;
        var rows = new List<TraceRow> { ToTraceRow(input.Root, 0, null) };
        var visited = new HashSet<string>
        {
            TraceUtils.BuildTraceCellKey(input.Root.WorksheetName, input.Root.RowIndex, input.Root.ColumnIndex)
        };
        var truncated = false;

        if (input.OnProgress != null)
        {
            await input.OnProgress(new TraceBuilderProgress
            {
                Rows = rows.ToArray(),
                Level = 0,
                IsFinal = input.MaxDepth == 0,
                Truncated = false
            }).ConfigureAwait(false);
        }

        var currentLevelCells = new List<TraceCellInfo> { input.Root };
        for (var level = 0; level < input.MaxDepth && currentLevelCells.Count > 0 && !truncated; level++)
        {
            var neighborLists = await input.GetAllNeighbors(currentLevelCells).ConfigureAwait(false);
            var nextLevelCells = new List<TraceCellInfo>();

            for (var i = 0; i < neighborLists.Count && !truncated; i++)
            {
                var parent = currentLevelCells[i];
                var neighbors = neighborLists[i] ?? Array.Empty<TraceCellInfo>();
                foreach (var neighbor in neighbors)
                {
                    var key = TraceUtils.BuildTraceCellKey(neighbor.WorksheetName, neighbor.RowIndex, neighbor.ColumnIndex);
                    if (!visited.Add(key)) continue;

                    rows.Add(ToTraceRow(neighbor, level + 1, parent.Address));
                    if (rows.Count >= maxRows)
                    {
                        truncated = true;
                        break;
                    }

                    nextLevelCells.Add(neighbor);
                }
            }

            currentLevelCells = nextLevelCells;

            if (input.OnProgress != null)
            {
                var willBeFinal = truncated || currentLevelCells.Count == 0 || level + 1 >= input.MaxDepth;
                await input.OnProgress(new TraceBuilderProgress
                {
                    Rows = rows.ToArray(),
                    Level = level + 1,
                    IsFinal = willBeFinal,
                    Truncated = truncated
                }).ConfigureAwait(false);
            }
        }

        return new TraceBuilderResult { Rows = rows, Truncated = truncated };
    }
}

public sealed class DependentEntry
{
    public string Address { get; set; } = "";
    public string Value { get; set; } = "";
    public int Count { get; set; }
}

public static class DependentsAggregator
{
    public static IReadOnlyList<DependentEntry> Aggregate(IEnumerable<TraceRow> rows, TraceDirection direction)
    {
        var counts = new Dictionary<string, (int Count, string Value)>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            if (row.Level == 0) continue;
            if (counts.TryGetValue(row.Address, out var existing))
            {
                counts[row.Address] = (existing.Count + 1, existing.Value);
            }
            else
            {
                counts[row.Address] = (1, row.Value);
            }
        }

        var result = new List<DependentEntry>();
        foreach (var pair in counts)
        {
            result.Add(new DependentEntry
            {
                Address = pair.Key,
                Value = pair.Value.Value,
                Count = pair.Value.Count
            });
        }

        result.Sort((a, b) => string.Compare(a.Address, b.Address, StringComparison.OrdinalIgnoreCase));
        return result;
    }
}
