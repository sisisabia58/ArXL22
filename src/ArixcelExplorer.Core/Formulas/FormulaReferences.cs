using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ArixcelExplorer.Core.Tracing;

namespace ArixcelExplorer.Core.Formulas;

public abstract class FormulaSegment
{
    public abstract string SegmentKind { get; }
}

public sealed class FormulaTextSegment : FormulaSegment
{
    public override string SegmentKind => "text";
    public string Text { get; set; } = "";
}

public sealed class FormulaReferenceSegment : FormulaSegment
{
    public override string SegmentKind => "reference";
    public string Text { get; set; } = "";
    public string? ReferenceKind { get; set; }
    public TraceTarget Target { get; set; } = new TraceAddressTargetWrapper();
}

public sealed class FormulaExternalReferenceSegment : FormulaSegment
{
    public override string SegmentKind => "external";
    public string Text { get; set; } = "";
    public string WorkbookName { get; set; } = "";
    public string WorksheetName { get; set; } = "";
    public string RangeAddress { get; set; } = "";
    public string Reason { get; set; } = "External workbook navigation requires the workbook to be open.";
}

public static class FormulaReferences
{
    private const string CellPattern = @"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}";
    private const string RangePattern = CellPattern + @"(?::" + CellPattern + ")?";
    private const string SheetPattern = @"'(?:[^']|'')+'|[A-Za-z_\\][A-Za-z0-9_.]*";
    private const string NamePattern = @"[A-Za-z_\\][A-Za-z0-9_.]*";
    private const string ExternalReason =
        "External workbook navigation is not available unless the workbook is open.";

    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "TRUE", "FALSE"
    };

    public static IReadOnlyList<FormulaSegment> TokenizeFormulaReferences(string formula, string contextWorksheetName)
    {
        if (string.IsNullOrEmpty(formula)) return Array.Empty<FormulaSegment>();

        var segments = new List<FormulaSegment>();
        var index = 0;
        while (index < formula.Length)
        {
            if (formula[index] == '"')
            {
                var end = StringLiteralEnd(formula, index);
                PushText(segments, formula.Substring(index, end - index));
                index = end;
                continue;
            }

            var external = TryExternalReference(formula, index);
            if (external != null)
            {
                segments.Add(external);
                index += external.Text.Length;
                continue;
            }

            var a1 = Regex.Match(formula.Substring(index),
                $"^(?:(?<sheet>{SheetPattern})!)?(?<range>{RangePattern})");
            if (a1.Success && a1.Groups["range"].Success)
            {
                var token = a1.Value;
                var sheetToken = a1.Groups["sheet"].Value;
                var rangeAddress = a1.Groups["range"].Value;
                var previous = index > 0 ? formula[index - 1] : '\0';
                var next = index + token.Length < formula.Length ? formula[index + token.Length] : '\0';
                var invalidBoundary =
                    IsIdentifierCharacter(previous) ||
                    IsIdentifierCharacter(next) ||
                    previous == '[' || previous == ']' ||
                    (previous == ':' && !string.IsNullOrEmpty(sheetToken)) ||
                    next == '[' || next == '(';

                if (!invalidBoundary && IsValidRangeAddress(rangeAddress))
                {
                    var worksheetName = string.IsNullOrEmpty(sheetToken)
                        ? contextWorksheetName
                        : UnquoteWorksheetName(sheetToken);
                    segments.Add(new FormulaReferenceSegment
                    {
                        Text = token,
                        Target = new TraceAddressTargetWrapper
                        {
                            Target = new TraceAddressTarget
                            {
                                Address = $"{QuoteWorksheetName(worksheetName)}!{rangeAddress}"
                            }
                        }
                    });
                    index += token.Length;
                    continue;
                }
            }

            var identifier = Regex.Match(formula.Substring(index), $"^(?<name>{NamePattern})");
            if (identifier.Success && index + identifier.Length < formula.Length && formula[index + identifier.Length] == '[')
            {
                var bracketStart = index + identifier.Length;
                var end = BracketExpressionEnd(formula, bracketStart);
                if (end > bracketStart)
                {
                    var token = formula.Substring(index, end - index);
                    var previous = index > 0 ? formula[index - 1] : '\0';
                    var target = BuildTableTarget(identifier.Groups["name"].Value, formula.Substring(bracketStart, end - bracketStart));
                    if (!IsIdentifierCharacter(previous) && target != null)
                    {
                        segments.Add(new FormulaReferenceSegment
                        {
                            Text = token,
                            ReferenceKind = "table",
                            Target = new TraceTableTargetWrapper { Target = target }
                        });
                    }
                    else
                    {
                        PushText(segments, token);
                    }

                    index = end;
                    continue;
                }
            }

            if (formula[index] == '[')
            {
                var end = BracketExpressionEnd(formula, index);
                if (end > index)
                {
                    PushText(segments, formula.Substring(index, end - index));
                    index = end;
                    continue;
                }
            }

            var named = Regex.Match(formula.Substring(index), $"^(?:(?<sheet>{SheetPattern})!)?(?<name>{NamePattern})");
            if (named.Success && named.Groups["name"].Success)
            {
                var token = named.Value;
                var sheetToken = named.Groups["sheet"].Value;
                var name = named.Groups["name"].Value;
                var previous = index > 0 ? formula[index - 1] : '\0';
                var next = index + token.Length < formula.Length ? formula[index + token.Length] : '\0';
                var worksheetName = string.IsNullOrEmpty(sheetToken) ? null : UnquoteWorksheetName(sheetToken);
                var validName =
                    !IsIdentifierCharacter(previous) &&
                    !IsIdentifierCharacter(next) &&
                    next != '(' && next != '[' &&
                    !ReservedNames.Contains(name) &&
                    !(worksheetName != null && IsExcelWorkbookName(worksheetName));

                if (validName)
                {
                    segments.Add(new FormulaReferenceSegment
                    {
                        Text = token,
                        ReferenceKind = "namedRange",
                        Target = new TraceNamedRangeTargetWrapper
                        {
                            Target = new TraceNamedRangeTarget
                            {
                                Name = name,
                                FormulaWorksheetName = contextWorksheetName,
                                WorksheetName = worksheetName
                            }
                        }
                    });
                    index += token.Length;
                    continue;
                }
            }

            PushText(segments, formula[index].ToString());
            index += 1;
        }

        return segments;
    }

    private static void PushText(List<FormulaSegment> segments, string text)
    {
        if (text.Length == 0) return;
        if (segments.Count > 0 && segments[segments.Count - 1] is FormulaTextSegment previous)
        {
            previous.Text += text;
            return;
        }

        segments.Add(new FormulaTextSegment { Text = text });
    }

    private static int StringLiteralEnd(string formula, int start)
    {
        var index = start + 1;
        while (index < formula.Length)
        {
            if (formula[index] != '"')
            {
                index += 1;
                continue;
            }

            if (index + 1 < formula.Length && formula[index + 1] == '"')
            {
                index += 2;
                continue;
            }

            return index + 1;
        }

        return formula.Length;
    }

    private static int BracketExpressionEnd(string formula, int start)
    {
        var depth = 0;
        for (var index = start; index < formula.Length; index++)
        {
            if (formula[index] == '[') depth += 1;
            else if (formula[index] == ']')
            {
                depth -= 1;
                if (depth == 0) return index + 1;
            }
        }

        return start;
    }

    private static FormulaExternalReferenceSegment? TryExternalReference(string formula, int index)
    {
        var remaining = formula.Substring(index);
        var match = Regex.Match(remaining,
            @"^'(?<sheet>(?:[^']|'')*?)\[(?<workbook>[^\]]+)\](?<sheet2>(?:[^']|'')+)'!(?<range>" + RangePattern + ")") ??
            Regex.Match(remaining, @"^\[(?<workbook>[^\]]+)\](?<sheet>[A-Za-z_][A-Za-z0-9_.]*)!(?<range>" + RangePattern + ")");
        if (!match.Success) return null;
        var range = match.Groups["range"].Value;
        if (!IsValidRangeAddress(range)) return null;
        var nextIndex = index + match.Length;
        var next = nextIndex < formula.Length ? formula[nextIndex] : '\0';
        if (IsIdentifierCharacter(next)) return null;

        var sheet = match.Groups["sheet"].Success ? match.Groups["sheet"].Value : match.Groups["sheet2"].Value;
        return new FormulaExternalReferenceSegment
        {
            Text = match.Value,
            WorkbookName = match.Groups["workbook"].Value,
            WorksheetName = sheet.Replace("''", "'"),
            RangeAddress = range,
            Reason = ExternalReason
        };
    }

    private static TraceTableTarget? BuildTableTarget(string tableName, string bracketText)
    {
        var inner = bracketText.Trim().TrimStart('[').TrimEnd(']').Trim();
        var sectionOnly = SectionFromToken(inner);
        if (sectionOnly.HasValue)
        {
            return new TraceTableTarget { TableName = tableName, Section = sectionOnly.Value };
        }

        if (inner.IndexOf('[') < 0 && inner.IndexOf(']') < 0 && !inner.StartsWith("@", StringComparison.Ordinal))
        {
            return new TraceTableTarget
            {
                TableName = tableName,
                Section = TraceTableSection.Data,
                ColumnStart = inner
            };
        }

        return null;
    }

    private static TraceTableSection? SectionFromToken(string token)
    {
        return token.Trim().ToLowerInvariant() switch
        {
            "#all" => TraceTableSection.All,
            "#data" => TraceTableSection.Data,
            "#headers" => TraceTableSection.Headers,
            "#totals" => TraceTableSection.Totals,
            _ => null
        };
    }

    private static bool IsIdentifierCharacter(char value) =>
        char.IsLetterOrDigit(value) || value == '.' || value == '_';

    private static bool IsValidCellAddress(string address)
    {
        var match = Regex.Match(address, @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$");
        if (!match.Success) return false;
        return TraceUtils.ColumnIndex(match.Groups[1].Value) <= 16_384 &&
               int.Parse(match.Groups[2].Value) <= 1_048_576;
    }

    private static bool IsValidRangeAddress(string address) =>
        address.Split(':') is var parts && Array.TrueForAll(parts, IsValidCellAddress);

    private static string UnquoteWorksheetName(string token)
    {
        if (token.StartsWith("'") && token.EndsWith("'"))
        {
            return token.Substring(1, token.Length - 2).Replace("''", "'");
        }

        return token;
    }

    private static string QuoteWorksheetName(string name)
    {
        if (Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_.]*$")) return name;
        return $"'{name.Replace("'", "''")}'";
    }

    private static bool IsExcelWorkbookName(string value) =>
        Regex.IsMatch(value, @"\.(xlsx|xlsm|xlsb|xls)$", RegexOptions.IgnoreCase);
}
