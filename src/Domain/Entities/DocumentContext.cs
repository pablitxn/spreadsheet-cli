using System.Collections.Generic;

namespace SpreadsheetCLI.Domain.Entities;

/// <summary>
/// Document context gathered from comprehensive traversal
/// </summary>
public sealed class DocumentContext
{
    public int TotalRows { get; set; }
    public int TotalColumns { get; set; }
    public Dictionary<string, SheetInfo> Sheets { get; set; } = new();
    public Dictionary<string, ColumnStats> ColumnStatistics { get; set; } = new();
    public List<DataPattern> DetectedPatterns { get; set; } = new();
    public List<Relationship> CrossSheetRelationships { get; set; } = new();
    
    // Additional properties for compatibility
    public DocumentMetadata Metadata { get; set; } = new();
    public List<Relationship> Relationships { get; set; } = new();
    public List<DataPattern> DataPatterns { get; set; } = new();
}