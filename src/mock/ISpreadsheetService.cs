using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace SpreadsheetCLI.Mock;

/// <summary>
/// Unified interface for all spreadsheet operations
/// Combines analysis, metadata extraction, and formula execution capabilities
/// </summary>
public interface ISpreadsheetService
{
    /// <summary>
    /// Analyzes a natural language query against spreadsheet data
    /// </summary>
    Task<QueryAnalysisResult> AnalyzeQueryAsync(
        string query,
        DocumentMetadata metadata,
        Worksheet worksheet,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Detects the document format using stratified sampling
    /// </summary>
    Task<DocumentFormat> DetectDocumentFormatAsync(
        Worksheet worksheet,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Extracts document metadata based on the detected format
    /// </summary>
    Task<DocumentMetadata> ExtractDocumentMetadataAsync(
        Worksheet worksheet,
        DocumentFormat format,
        List<HeaderInfo> headers,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Extracts headers from a worksheet
    /// </summary>
    List<HeaderInfo> ExtractHeaders(Worksheet worksheet);

    /// <summary>
    /// Creates a dynamic spreadsheet from execution plan data
    /// </summary>
    Task<DynamicSpreadsheetResult> CreateDynamicSpreadsheetAsync(
        ExecutionPlanDto executionPlan,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes a formula on a workbook and returns the result
    /// </summary>
    Task<FormulaExecutionResult> ExecuteFormulaAsync(
        Workbook workbook,
        ExecutionPlanDto executionPlan,
        CancellationToken cancellationToken = default);
}

#region DTOs and Supporting Types

/// <summary>
/// Result from query analysis
/// </summary>
public sealed class QueryAnalysisResult
{
    public List<string> ColumnsNeeded { get; set; } = new();
    public List<FilterCriteria> Filters { get; set; } = new();
    public string AggregationType { get; set; } = "";
    public string? GroupBy { get; set; }
    public bool RequiresCalculation { get; set; }
    public List<string> CalculationSteps { get; set; } = new();
    public bool RequiresFullDataset { get; set; }
    public string UserIntentWithContext { get; set; } = "";
    public string Artifact { get; set; } = "";
}

/// <summary>
/// Document metadata information
/// </summary>
public sealed class DocumentMetadata
{
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public DocumentFormat Format { get; set; }
    public long FileSize { get; set; }
    public int SheetCount { get; set; }
    public DateTime CreatedDate { get; set; }
    public DateTime ModifiedDate { get; set; }
    public Dictionary<string, string> Properties { get; set; } = new();
    public int TotalRows { get; set; }
    public int TotalColumns { get; set; }
    public int DataStartRow { get; set; }
    public int DataRowCount { get; set; }
    public List<string> Headers { get; set; } = new();
    public Dictionary<string, string> DataTypes { get; set; } = new();
}

/// <summary>
/// Document format enumeration
/// </summary>
public enum DocumentFormat
{
    Unknown,
    Columnar,
    RowBased,
    Nested,
    Matrix,
    Mixed,
    Excel,
    CSV,
    TSV,
    Other
}

/// <summary>
/// Header information
/// </summary>
public sealed class HeaderInfo
{
    public string Name { get; }
    public int RowIndex { get; }

    public HeaderInfo(string name, int rowIndex)
    {
        Name = name;
        RowIndex = rowIndex;
    }
}

/// <summary>
/// Filter criteria for data queries
/// </summary>
public sealed class FilterCriteria
{
    public string Column { get; set; } = "";
    public string Operator { get; set; } = "";
    public string Value { get; set; } = "";
}

/// <summary>
/// Execution plan DTO
/// </summary>
public sealed class ExecutionPlanDto
{
    public bool NeedRunFormula { get; set; }
    public List<List<object>>? ArtifactsFormatted { get; set; }
    public string Formula { get; set; } = "";
    public string SimpleAnswer { get; set; } = "";
    public string Reasoning { get; set; } = "";
    public string? MachineAnswer { get; set; }
    public string? HumanExplanation { get; set; }
}

/// <summary>
/// Result from dynamic spreadsheet creation
/// </summary>
public sealed class DynamicSpreadsheetResult
{
    public Workbook Workbook { get; set; } = null!;
    public Worksheet Worksheet { get; set; } = null!;
    public int DataRows { get; set; }
    public int DataColumns { get; set; }
    public List<CellAssignment> CellAssignments { get; set; } = new();
}

/// <summary>
/// Cell assignment tracking
/// </summary>
public sealed class CellAssignment
{
    public string CellReference { get; set; } = "";
    public int Row { get; set; }
    public int Column { get; set; }
    public object? OriginalValue { get; set; }
    public string OriginalType { get; set; } = "";
    public object? AssignedValue { get; set; }
    public string AssignedType { get; set; } = "";
    public bool IsJsonElement { get; set; }
}

/// <summary>
/// Formula execution result
/// </summary>
public sealed class FormulaExecutionResult
{
    public bool Success { get; set; }
    public object? Value { get; set; }
    public string StringValue { get; set; } = "";
    public string? Error { get; set; }
    public string FormulaCellReference { get; set; } = "";
    public Dictionary<string, object> DebugInfo { get; set; } = new();
}

#endregion