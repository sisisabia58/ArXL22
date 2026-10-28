using System;

namespace ArixcelExplorer.Core.Formulas;

public static class FunctionArgInfo
{
    public static string LabelFor(string functionName, int argumentIndex)
    {
        var name = (functionName ?? "").ToUpperInvariant();
        return name switch
        {
            "IF" => argumentIndex switch
            {
                0 => "logical_test",
                1 => "value_if_true",
                2 => "value_if_false",
                _ => ""
            },
            "SUMIFS" => argumentIndex == 0 ? "sum_range" : CriteriaPairLabel(argumentIndex),
            "AVERAGEIFS" => argumentIndex == 0 ? "average_range" : CriteriaPairLabel(argumentIndex),
            "MINIFS" => argumentIndex == 0 ? "min_range" : CriteriaPairLabel(argumentIndex),
            "MAXIFS" => argumentIndex == 0 ? "max_range" : CriteriaPairLabel(argumentIndex),
            "COUNTIFS" => CountIfsLabel(argumentIndex),
            "SUMIF" => argumentIndex switch
            {
                0 => "range",
                1 => "criteria",
                2 => "sum_range",
                _ => ""
            },
            "COUNTIF" => argumentIndex switch
            {
                0 => "range",
                1 => "criteria",
                _ => ""
            },
            "AVERAGEIF" => argumentIndex switch
            {
                0 => "range",
                1 => "criteria",
                2 => "average_range",
                _ => ""
            },
            "INDEX" => argumentIndex switch
            {
                0 => "array",
                1 => "row_num",
                2 => "column_num",
                _ => ""
            },
            "OFFSET" => argumentIndex switch
            {
                0 => "reference",
                1 => "rows",
                2 => "cols",
                3 => "height",
                4 => "width",
                _ => ""
            },
            _ => ""
        };
    }

    /// <summary>
    /// SUMIFS-style pairs start at argument 1: criteria_range1, criteria1, criteria_range2, criteria2.
    /// </summary>
    private static string CriteriaPairLabel(int argumentIndex) =>
        argumentIndex % 2 == 1
            ? "criteria_range" + ((argumentIndex + 1) / 2)
            : "criteria" + (argumentIndex / 2);

    private static string CountIfsLabel(int argumentIndex) =>
        argumentIndex % 2 == 0
            ? "criteria_range" + (argumentIndex / 2 + 1)
            : "criteria" + (argumentIndex / 2 + 1);
}
