using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using SpreadsheetCLI.Domain.Enums;
using SpreadsheetCLI.Domain.Entities;
using SpreadsheetCLI.Domain.ValueObjects;
using SpreadsheetCLI.Application.DTOs;

namespace SpreadsheetCLI.Application.Interfaces.Spreadsheet;

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