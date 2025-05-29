namespace SpreadsheetCLI.Domain.ValueObjects;

using System;

/// <summary>
/// Filter criteria for data filtering
/// </summary>
public sealed class FilterCriteria
{
    public string Column { get; set; }
    public string Operator { get; set; }
    public string Value { get; set; }

    public FilterCriteria()
    {
        Column = "";
        Operator = "";
        Value = "";
    }

    public FilterCriteria(string column, string @operator, string value)
    {
        Column = column ?? "";
        Operator = @operator ?? "";
        Value = value ?? "";
    }

    public override bool Equals(object? obj)
    {
        if (obj is not FilterCriteria other)
            return false;

        return Column == other.Column &&
               Operator == other.Operator &&
               Value == other.Value;
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Column, Operator, Value);
    }
}