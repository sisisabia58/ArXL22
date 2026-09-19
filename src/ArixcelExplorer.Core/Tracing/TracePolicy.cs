using System;

namespace ArixcelExplorer.Core.Tracing;

public sealed class TraceLoadingPolicy
{
    public int RangeSummaryThreshold { get; set; }
    public int RangePageSize { get; set; }
    public int VisibleNodeLimit { get; set; }
    public int HardNodeLimit { get; set; }
    public int PrefetchMaxCandidates { get; set; }
    public int PrefetchMaxCells { get; set; }
}

public static class TracePolicy
{
    public const int DefaultTraceRangeSummaryThreshold = 50;
    public const int DefaultTraceRangePageSize = 100;
    public const int DefaultTraceVisibleNodeLimit = 500;
    public const int DefaultTracePrefetchMaxCandidates = 20;
    public const int DefaultTracePrefetchMaxCells = 50;

    public static TraceLoadingPolicy BuildTraceLoadingPolicy(int? safetyLimit = null, int? visibleNodeLimit = null)
    {
        var hardNodeLimit = Math.Min(
            TraceUtils.MaxTraceSafetyLimit,
            TraceUtils.SanitizeTraceSafetyLimit(safetyLimit));

        return new TraceLoadingPolicy
        {
            RangeSummaryThreshold = DefaultTraceRangeSummaryThreshold,
            RangePageSize = DefaultTraceRangePageSize,
            VisibleNodeLimit = Math.Min(visibleNodeLimit ?? DefaultTraceVisibleNodeLimit, hardNodeLimit),
            HardNodeLimit = hardNodeLimit,
            PrefetchMaxCandidates = DefaultTracePrefetchMaxCandidates,
            PrefetchMaxCells = DefaultTracePrefetchMaxCells
        };
    }

    public static bool ShouldSummarizeTraceRange(int cellCount, TraceLoadingPolicy policy)
    {
        return cellCount > policy.RangeSummaryThreshold;
    }

    public static (int Start, int Count, bool HasMore) GetRangePageBounds(
        int cellCount,
        int page,
        TraceLoadingPolicy policy)
    {
        var normalizedPage = page >= 0 ? page : 0;
        var start = Math.Min(cellCount, normalizedPage * policy.RangePageSize);
        var count = Math.Min(policy.RangePageSize, Math.Max(0, cellCount - start));
        return (start, count, start + count < cellCount);
    }
}
