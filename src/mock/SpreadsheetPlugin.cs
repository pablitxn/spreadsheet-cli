using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;
using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel;

namespace SpreadsheetCLI.Mock;

/// <summary>
/// Semantic Kernel plugin for spreadsheet operations
/// Uses the unified SpreadsheetService for all functionality
/// </summary>
public sealed class SpreadsheetPlugin
{
    private readonly ILogger<SpreadsheetPlugin> _logger;
    private readonly ISpreadsheetService _spreadsheetService;

    public SpreadsheetPlugin(
        ILogger<SpreadsheetPlugin> logger,
        ISpreadsheetService spreadsheetService)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _spreadsheetService = spreadsheetService ?? throw new ArgumentNullException(nameof(spreadsheetService));
    }

    [KernelFunction("query_spreadsheet")]
    [Description("Queries Excel data using natural language with high accuracy through sandbox execution")]
    public async Task<string> QuerySpreadsheetAsync(
        [Description("Full path of the workbook")]
        string filePath,
        [Description("Natural language query")]
        string query,
        [Description("Sheet name (optional)")] string sheetName = "")
    {
        _logger.LogInformation("QuerySpreadsheet: {File} - '{Query}'", filePath, query);

        try
        {
            // Step 1: Load workbook
            using var workbook = await LoadWorkbookAsync(filePath);
            var sourceSheet = GetWorksheet(workbook, sheetName);

            // Step 2: Detect document format
            var format = await _spreadsheetService.DetectDocumentFormatAsync(sourceSheet);

            // Step 3: Validate format
            if (format != DocumentFormat.Columnar)
            {
                throw new InvalidOperationException(
                    $"Unsupported spreadsheet format: {format}. Only columnar format is currently supported.");
            }

            // Step 4: Extract headers
            var headers = _spreadsheetService.ExtractHeaders(sourceSheet);
            _logger.LogInformation("Extracted {HeaderCount} headers from sheet {SheetName}", 
                headers.Count, sourceSheet.Name);

            // Step 5: Extract metadata
            var metadata = await _spreadsheetService.ExtractDocumentMetadataAsync(sourceSheet, format, headers);
            _logger.LogInformation(
                "Metadata extracted - Total rows: {TotalRows}, Data rows: {DataRows}, Columns: {Columns}",
                metadata.TotalRows, metadata.DataRowCount, metadata.TotalColumns);

            // Step 6: Analyze the query
            var analysisResult = await _spreadsheetService.AnalyzeQueryAsync(query, metadata, sourceSheet);

            // Extract execution plan
            var executionPlan = ExtractExecutionPlan(analysisResult);

            _logger.LogInformation(
                "Query analysis complete - Need formula: {NeedFormula}, Formula: {Formula}",
                executionPlan.NeedRunFormula, executionPlan.Formula);

            // Check if we need to run a formula
            if (!executionPlan.NeedRunFormula)
            {
                return CreateSimpleResult(query, executionPlan, metadata);
            }

            // Step 7: Create dynamic spreadsheet
            var dynamicSpreadsheetResult = await _spreadsheetService.CreateDynamicSpreadsheetAsync(executionPlan);
            var dynamicWorkbook = dynamicSpreadsheetResult.Workbook;

            _logger.LogInformation(
                "Dynamic spreadsheet created - Rows: {Rows}, Columns: {Columns}",
                dynamicSpreadsheetResult.DataRows, dynamicSpreadsheetResult.DataColumns);

            try
            {
                // Step 8: Execute the formula
                var formulaExecution = await _spreadsheetService.ExecuteFormulaAsync(dynamicWorkbook, executionPlan);

                _logger.LogInformation(
                    "Formula executed - Success: {Success}, Result: {Result}",
                    formulaExecution.Success, formulaExecution.StringValue);

                // Step 9: Create final result
                return CreateFormulaResult(query, executionPlan, metadata, formulaExecution);
            }
            finally
            {
                // Cleanup dynamic workbook
                dynamicWorkbook.Dispose();
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "QuerySpreadsheet failed");
            return CreateErrorResult(query, ex);
        }
    }

    #region Private Helper Methods

    private async Task<Workbook> LoadWorkbookAsync(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Spreadsheet file not found: {path}");
        }

        var stream = File.OpenRead(path);
        return new Workbook(stream);
    }

    private Worksheet GetWorksheet(Workbook workbook, string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return workbook.Worksheets[0];

        return workbook.Worksheets.FirstOrDefault(w =>
            w.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) ?? workbook.Worksheets[0];
    }

    private ExecutionPlanDto ExtractExecutionPlan(QueryAnalysisResult analysisResult)
    {
        try
        {
            using var doc = JsonDocument.Parse(analysisResult.Artifact);
            var executionPlanElement = doc.RootElement.GetProperty("ExecutionPlan");
            return JsonSerializer.Deserialize<ExecutionPlanDto>(executionPlanElement.GetRawText())
                   ?? new ExecutionPlanDto();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to extract execution plan from artifact");
            return new ExecutionPlanDto();
        }
    }

    private string CreateSimpleResult(string query, ExecutionPlanDto executionPlan, DocumentMetadata metadata)
    {
        var result = new
        {
            Success = true,
            Query = query,
            Answer = executionPlan.MachineAnswer ?? executionPlan.SimpleAnswer,
            Reasoning = executionPlan.Reasoning,
            HumanExplanation = executionPlan.HumanExplanation,
            RequiredCalculation = false,
            DatasetContext = CreateDatasetContext(metadata)
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    private string CreateFormulaResult(
        string query,
        ExecutionPlanDto executionPlan,
        DocumentMetadata metadata,
        FormulaExecutionResult formulaExecution)
    {
        var result = new
        {
            Success = true,
            Query = query,
            Answer = executionPlan.MachineAnswer ?? formulaExecution.StringValue,
            Formula = executionPlan.Formula,
            Reasoning = executionPlan.Reasoning,
            HumanExplanation = executionPlan.HumanExplanation,
            RequiredCalculation = true,
            DatasetContext = CreateDatasetContext(metadata),
            DataUsed = new
            {
                Rows = executionPlan.ArtifactsFormatted?.Count ?? 0,
                Columns = executionPlan.ArtifactsFormatted?.FirstOrDefault()?.Count ?? 0,
                Headers = executionPlan.ArtifactsFormatted?.FirstOrDefault()
            }
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    private string CreateErrorResult(string query, Exception ex)
    {
        var result = new
        {
            Success = false,
            Error = ex.Message,
            Query = query
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    private object CreateDatasetContext(DocumentMetadata metadata)
    {
        return new
        {
            HeaderRowIndex = metadata.DataStartRow - 1,
            DataStartRow = metadata.DataStartRow,
            DataEndRow = metadata.TotalRows - 1,
            TotalDataRows = metadata.DataRowCount,
            Explanation = $"The dataset has headers at row {metadata.DataStartRow - 1} and contains {metadata.DataRowCount} data rows (from row {metadata.DataStartRow} to {metadata.TotalRows - 1})"
        };
    }

    #endregion
}