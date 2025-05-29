namespace SpreadsheetCLI.Application.DTOs;

/// <summary>
/// Represents a cell assignment in the dynamic spreadsheet
/// </summary>
public sealed class CellAssignment
{
    public string CellReference { get; set; } = "";
    public int Row { get; set; }
    public int Column { get; set; }
    public object? OriginalValue { get; set; }
    public object? AssignedValue { get; set; }
    public string OriginalType { get; set; } = "";
    public string AssignedType { get; set; } = "";
    public bool IsJsonElement { get; set; }
}