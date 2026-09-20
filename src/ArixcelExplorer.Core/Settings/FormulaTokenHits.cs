using System;
using System.Collections.Generic;
using ArixcelExplorer.Core.Formulas;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.Core.Settings;

public sealed class FormulaTokenSpan
{
    public int Start { get; set; }
    public int Length { get; set; }
    public string Text { get; set; } = "";
    public string? Address { get; set; }
}

public static class FormulaTokenHits
{
    public static IReadOnlyList<FormulaTokenSpan> Spans(string formula, string contextWorksheetName)
    {
        var result = new List<FormulaTokenSpan>();
        if (string.IsNullOrEmpty(formula)) return result;

        var sheet = string.IsNullOrWhiteSpace(contextWorksheetName) ? "Sheet1" : contextWorksheetName;
        var index = 0;
        foreach (var segment in FormulaReferences.TokenizeFormulaReferences(formula, sheet))
        {
            var text = segment switch
            {
                FormulaReferenceSegment reference => reference.Text,
                FormulaExternalReferenceSegment external => external.Text,
                FormulaTextSegment plain => plain.Text,
                _ => ""
            };
            if (segment is FormulaReferenceSegment referenceSegment)
            {
                string? address = null;
                if (referenceSegment.Target is TraceAddressTargetWrapper wrapper)
                {
                    address = NormalizeAddress(wrapper.Target?.Address);
                }

                result.Add(new FormulaTokenSpan
                {
                    Start = index,
                    Length = text.Length,
                    Text = text,
                    Address = address
                });
            }

            index += text.Length;
        }

        return result;
    }

    public static FormulaTokenSpan? HitTest(string formula, string contextWorksheetName, int charIndex)
    {
        if (charIndex < 0) return null;
        foreach (var span in Spans(formula, contextWorksheetName))
        {
            if (charIndex >= span.Start && charIndex < span.Start + span.Length)
            {
                return span;
            }
        }

        return null;
    }

    private static string? NormalizeAddress(string? address)
    {
        if (address is null || string.IsNullOrWhiteSpace(address)) return address;
        var parsed = TraceUtils.ParseWorksheetScopedAddress(address);
        if (parsed == null) return address;
        return $"'{parsed.WorksheetName}'!{parsed.RangeAddress}";
    }
}
