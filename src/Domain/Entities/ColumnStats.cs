using System.Collections.Generic;
using SpreadsheetCLI.Domain.Enums;

namespace SpreadsheetCLI.Domain.Entities;

/// <summary>
/// Statistics for a column
/// </summary>
public sealed class ColumnStats
{
    public string ColumnName { get; set; } = "";
    public ColumnType DataType { get; set; }
    public int NonNullCount { get; set; }
    public int UniqueValueCount { get; set; }
    public object? MinValue { get; set; }
    public object? MaxValue { get; set; }
    public double? Average { get; set; }
    public List<string> SampleValues { get; set; } = new();
    
    // Additional properties for compatibility
    public int TotalCount { get; set; }
    public int UniqueCount { get; set; }
    public int NullCount { get; set; }
    public double? Sum { get; set; }
    public double? Min { get; set; }
    public double? Max { get; set; }
}