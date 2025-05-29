using System.Collections.Generic;
using Aspose.Cells;

namespace SpreadsheetCLI.Application.DTOs;

/// <summary>
/// Result of creating a dynamic spreadsheet
/// </summary>
public sealed class DynamicSpreadsheetResult
{
    public Workbook Workbook { get; set; } = null!;
    public Worksheet Worksheet { get; set; } = null!;
    public int DataRows { get; set; }
    public int DataColumns { get; set; }
    public List<CellAssignment> CellAssignments { get; set; } = new();
}