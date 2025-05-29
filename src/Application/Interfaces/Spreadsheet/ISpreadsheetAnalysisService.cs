using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace SpreadsheetCLI.Core.Application.Interfaces.Spreadsheet;

/// <summary>
/// Service for analyzing spreadsheet data and extracting contextual information
/// </summary>
public interface ISpreadsheetAnalysisService
{
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
    /// Analyzes a query to determine required columns, filters, and operations
    /// </summary>
    Task<QueryAnalysisResult> AnalyzeQueryAsync(
        string query,
        DocumentMetadata metadata,
        Worksheet worksheet,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Extracts headers from a worksheet
    /// </summary>
    List<HeaderInfo> ExtractHeaders(Worksheet worksheet);

    /// <summary>
    /// Counts data rows in a worksheet
    /// </summary>
    int CountDataRows(Worksheet worksheet);

    /// <summary>
    /// Gets the data range of a worksheet
    /// </summary>
    (int FirstRow, int LastRow) GetDataRange(Worksheet worksheet);

    /// <summary>
    /// Checks if a row matches the given filters
    /// </summary>
    bool RowMatchesFilters(
        Worksheet worksheet,
        int row,
        List<HeaderInfo> headers,
        List<FilterCriteria> filters);

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

/// <summary>
/// Document format types that can be detected
/// </summary>
public enum DocumentFormat
{
    Columnar,
    RowBased,
    Nested,
    Matrix,
    Mixed,
    Unknown
}

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
}

/// <summary>
/// Result of query analysis
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
    public List<object> ContextSnapshots { get; set; } = new();
}

/// <summary>
/// Filter criteria for data filtering
/// </summary>
public sealed class FilterCriteria
{
    public string Column { get; set; } = "";
    public string Operator { get; set; } = "";
    public string Value { get; set; } = "";
}

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
}

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
}

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
}

/// <summary>
/// Column data type
/// </summary>
public enum ColumnType
{
    Text,
    Numeric,
    Date,
    Boolean,
    Formula,
    Mixed
}

/// <summary>
/// Detected data pattern
/// </summary>
public sealed class DataPattern
{
    public string PatternType { get; set; } = "";
    public string Description { get; set; } = "";
    public string Location { get; set; } = "";
    public double Confidence { get; set; }
}

/// <summary>
/// Cross-sheet relationship
/// </summary>
public sealed class Relationship
{
    public string SourceSheet { get; set; } = "";
    public string SourceColumn { get; set; } = "";
    public string TargetSheet { get; set; } = "";
    public string TargetColumn { get; set; } = "";
    public string RelationType { get; set; } = "";
}

/// <summary>
/// Header information including name and row index
/// </summary>
public sealed class HeaderInfo
{
    public string Name { get; set; } = "";
    public int RowIndex { get; set; }

    public HeaderInfo()
    {
    }

    public HeaderInfo(string name, int rowIndex)
    {
        Name = name;
        RowIndex = rowIndex;
    }
}

/// <summary>
/// DTO for the execution plan from analysis
/// </summary>
public sealed class ExecutionPlanDto
{
    public bool NeedRunFormula { get; set; }
    public List<List<object>>? ArtifactsFormatted { get; set; }
    public string Formula { get; set; } = "";
    public string SimpleAnswer { get; set; } = "";
    public string Reasoning { get; set; } = "";
}

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

/// <summary>
/// Result of formula execution
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