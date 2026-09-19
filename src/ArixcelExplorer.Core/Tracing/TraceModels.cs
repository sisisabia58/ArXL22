using System.Collections.Generic;

namespace ArixcelExplorer.Core.Tracing;

public enum TraceNodeLoadState
{
    Unloaded,
    Loading,
    Loaded,
    Leaf,
    Error
}

public class TraceCellData
{
    public const string KindName = "cell";
    public string Key { get; set; } = "";
    public string WorksheetName { get; set; } = "";
    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
    public string Address { get; set; } = "";
    public string Value { get; set; } = "";
    public string Formula { get; set; } = "";
}

public class TraceRangeData
{
    public const string KindName = "range";
    public string Key { get; set; } = "";
    public string WorksheetName { get; set; } = "";
    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public int CellCount { get; set; }
    public string Address { get; set; } = "";
}

public abstract class TraceNodeData
{
    public abstract string Kind { get; }
}

public sealed class TraceCellNodeData : TraceNodeData
{
    public override string Kind => TraceCellData.KindName;
    public TraceCellData Cell { get; set; } = new();
}

public sealed class TraceRangeNodeData : TraceNodeData
{
    public override string Kind => TraceRangeData.KindName;
    public TraceRangeData Range { get; set; } = new();
}

public sealed class TraceNodePlacement
{
    public string Id { get; set; } = "";
    public string? ParentId { get; set; }
    public int Level { get; set; }
    public TraceNodeLoadState LoadState { get; set; }
}

public sealed class TraceCellNode : TraceCellData, ITraceNode
{
    public string Id { get; set; } = "";
    public string? ParentId { get; set; }
    public int Level { get; set; }
    public TraceNodeLoadState LoadState { get; set; }
    public string NodeKind => KindName;
}

public sealed class TraceRangeNode : TraceRangeData, ITraceNode
{
    public string Id { get; set; } = "";
    public string? ParentId { get; set; }
    public int Level { get; set; }
    public TraceNodeLoadState LoadState { get; set; }
    public string NodeKind => KindName;
}

public sealed class TraceReferenceNode : ITraceNode
{
    public string NodeKind => "reference";
    public string Id { get; set; } = "";
    public string? ParentId { get; set; }
    public int Level { get; set; }
    public TraceNodeLoadState LoadState { get; set; }
    public string Key { get; set; } = "";
    public string TargetId { get; set; } = "";
    public string WorksheetName { get; set; } = "";
    public string Address { get; set; } = "";
    public string Value { get; set; } = "";
    public string Formula { get; set; } = "";
}

public interface ITraceNode
{
    string Id { get; }
    string? ParentId { get; }
    int Level { get; }
    TraceNodeLoadState LoadState { get; }
    string NodeKind { get; }
}

public sealed class TraceAddressTarget
{
    public string Address { get; set; } = "";
}

public sealed class TraceNamedRangeTarget
{
    public string Name { get; set; } = "";
    public string FormulaWorksheetName { get; set; } = "";
    public string? WorksheetName { get; set; }
}

public enum TraceTableSection
{
    All,
    Data,
    Headers,
    Totals
}

public sealed class TraceTableTarget
{
    public string TableName { get; set; } = "";
    public TraceTableSection Section { get; set; }
    public string? ColumnStart { get; set; }
    public string? ColumnEnd { get; set; }
}

public abstract class TraceTarget
{
    public abstract string TargetKind { get; }
}

public sealed class TraceAddressTargetWrapper : TraceTarget
{
    public override string TargetKind => "address";
    public TraceAddressTarget Target { get; set; } = new();
}

public sealed class TraceNamedRangeTargetWrapper : TraceTarget
{
    public override string TargetKind => "namedRange";
    public TraceNamedRangeTarget Target { get; set; } = new();
}

public sealed class TraceTableTargetWrapper : TraceTarget
{
    public override string TargetKind => "table";
    public TraceTableTarget Target { get; set; } = new();
}

public sealed class DirectNeighborResult
{
    public IReadOnlyList<TraceNodeData> Nodes { get; set; } = new List<TraceNodeData>();
    public bool Truncated { get; set; }
    public int MaterializedCellCount { get; set; }
}

public sealed class TraceRangePageResult
{
    public IReadOnlyList<TraceCellData> Nodes { get; set; } = new List<TraceCellData>();
    public int Page { get; set; }
    public bool HasMore { get; set; }
    public bool Truncated { get; set; }
}

public static class TraceGraphHelpers
{
    public static string BuildTraceRangeKey(
        string worksheetName,
        int rowIndex,
        int columnIndex,
        int rowCount,
        int columnCount)
    {
        return $"{worksheetName}!R{rowIndex}C{columnIndex}:{rowCount}x{columnCount}";
    }

    public static bool IsTraceNodeExpandable(ITraceNode node)
    {
        return node.NodeKind != "reference" && node.LoadState != TraceNodeLoadState.Leaf;
    }
}
