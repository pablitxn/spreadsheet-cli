using System.Collections.Generic;
using SpreadsheetCLI.Domain.Enums;
using SpreadsheetCLI.Domain.ValueObjects;

namespace SpreadsheetCLI.Domain.Entities;

/// <summary>
/// Information about a single sheet
/// </summary>
public sealed class SheetInfo
{
    public string Name { get; set; } = "";
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public List<HeaderInfo> Headers { get; set; } = new();
    public Dictionary<string, ColumnType> ColumnTypes { get; set; } = new();
    public List<string> FormulaCells { get; set; } = new();
    public bool IsHidden { get; set; }
}