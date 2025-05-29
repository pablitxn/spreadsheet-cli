using System;
using System.Collections.Generic;
using SpreadsheetCLI.Domain.Enums;

namespace SpreadsheetCLI.Domain.Entities;

/// <summary>
/// Document metadata containing structure and type information
/// </summary>
public sealed class DocumentMetadata
{
    public DocumentFormat Format { get; set; }
    public int TotalRows { get; set; }
    public int TotalColumns { get; set; }
    public List<string> Headers { get; set; } = new();
    public Dictionary<string, string> DataTypes { get; set; } = new();
    public int DataStartRow { get; set; } = 1;
    public int DataRowCount { get; set; }
    
    // File metadata properties
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public long FileSize { get; set; }
    public int SheetCount { get; set; }
    public DateTime CreatedDate { get; set; }
    public DateTime ModifiedDate { get; set; }
    public Dictionary<string, string> Properties { get; set; } = new();
}