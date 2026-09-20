using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ArixcelExplorer.Core.Tracing;
using Xunit;

namespace ArixcelExplorer.Core.Tests.Tracing;

public sealed class TraceBuilderTests
{
    [Fact]
    public async Task BuildAsync_visits_neighbors_in_bfs_order()
    {
        var root = new TraceCellInfo
        {
            WorksheetName = "Sheet1",
            RowIndex = 0,
            ColumnIndex = 0,
            Address = "'Sheet1'!A1",
            Value = 10,
            Formula = "=B1+C1"
        };

        var neighbors = new Dictionary<string, TraceCellInfo[]>
        {
            ["Sheet1!R0C0"] = new[]
            {
                Cell("Sheet1", 0, 1, "B1"),
                Cell("Sheet1", 0, 2, "C1")
            },
            ["Sheet1!R0C1"] = new[] { Cell("Sheet1", 1, 1, "B2") }
        };

        var result = await TraceBuilder.BuildAsync(new TraceBuilderInput
        {
            Root = root,
            MaxDepth = 2,
            GetAllNeighbors = cells =>
            {
                var lists = new List<IReadOnlyList<TraceCellInfo>>();
                foreach (var cell in cells)
                {
                    var key = TraceUtils.BuildTraceCellKey(cell.WorksheetName, cell.RowIndex, cell.ColumnIndex);
                    lists.Add(neighbors.TryGetValue(key, out var list) ? list : System.Array.Empty<TraceCellInfo>());
                }

                return Task.FromResult<IReadOnlyList<IReadOnlyList<TraceCellInfo>>>(lists);
            }
        });

        Assert.Equal(4, result.Rows.Count);
        Assert.Contains(result.Rows, r => r.Address.Contains("B2"));
    }

    [Fact]
    public void DependentsAggregator_counts_duplicate_addresses()
    {
        var rows = new[]
        {
            new TraceRow { Level = 0, Address = "'S'!A1", Value = "1" },
            new TraceRow { Level = 1, Address = "'S'!B1", Value = "2" },
            new TraceRow { Level = 1, Address = "'S'!B1", Value = "2" }
        };

        var aggregated = DependentsAggregator.Aggregate(rows, TraceDirection.Dependents);
        Assert.Single(aggregated);
        Assert.Equal(2, aggregated[0].Count);
    }

    [Fact]
    public void PrependOrigin_puts_origin_first_with_count_zero()
    {
        var dependents = new[]
        {
            new DependentEntry { Address = "'Sheet1'!B2", Value = "10", Count = 2 },
            new DependentEntry { Address = "'Sheet1'!C3", Value = "20", Count = 1 }
        };

        var rows = DependentsAggregator.PrependOrigin("'Sheet1'!A1", "99", dependents);

        Assert.Equal(3, rows.Count);
        Assert.Equal("'Sheet1'!A1", rows[0].Address);
        Assert.Equal("99", rows[0].Value);
        Assert.Equal(0, rows[0].Count);
        Assert.Equal("'Sheet1'!B2", rows[1].Address);
        Assert.Equal(2, rows[1].Count);
        Assert.Equal("A1", DependentsAggregator.ElementAddress("'Sheet1'!A1"));
        Assert.Equal("G15", DependentsAggregator.ElementAddress("'Summary (Instit)'!G15"));
    }

    [Fact]
    public void Aggregate_with_only_origin_row_returns_empty_dependents()
    {
        var rows = new[]
        {
            new TraceRow { Level = 0, Address = "'Sheet1'!A1", Value = "1" }
        };

        var entries = DependentsAggregator.Aggregate(rows, TraceDirection.Dependents);

        Assert.Empty(entries);
        var withOrigin = DependentsAggregator.PrependOrigin("'Sheet1'!A1", "1", entries);
        Assert.Single(withOrigin);
        Assert.Equal("'Sheet1'!A1", withOrigin[0].Address);
        Assert.Equal(0, withOrigin[0].Count);
    }

    [Fact]
    public async Task BuildAsync_returns_origin_only_when_neighbors_throw()
    {
        var root = Cell("Sheet1", 0, 0, "A1");
        var result = await TraceBuilder.BuildAsync(new TraceBuilderInput
        {
            Root = root,
            MaxDepth = 1,
            GetAllNeighbors = _ => throw new InvalidOperationException("Excel has no dependents")
        });

        Assert.Single(result.Rows);
        Assert.Equal("'Sheet1'!A1", result.Rows[0].Address);
        Assert.Empty(DependentsAggregator.Aggregate(result.Rows, TraceDirection.Dependents));
    }

    [Fact]
    public void Build_sync_returns_origin_only_when_neighbors_throw()
    {
        var root = Cell("Sheet1", 0, 0, "A1");
        var result = TraceBuilder.Build(root, 1, 100, _ => throw new InvalidOperationException("Excel has no dependents"));

        Assert.Single(result.Rows);
        Assert.Equal("'Sheet1'!A1", result.Rows[0].Address);
    }

    private static TraceCellInfo Cell(string sheet, int row, int col, string addr) => new()
    {
        WorksheetName = sheet,
        RowIndex = row,
        ColumnIndex = col,
        Address = $"'{sheet}'!{addr}",
        Value = 1
    };
}
