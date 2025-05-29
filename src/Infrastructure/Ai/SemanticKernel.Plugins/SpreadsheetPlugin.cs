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
using SpreadsheetCLI.Core.Application.Interfaces;
using SpreadsheetCLI.Core.Application.Interfaces.Spreadsheet;

namespace SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Plugins;

/// <summary>
/// SpreadsheetPluginV3 (Refactored): Enhanced Excel plugin using decoupled services for better maintainability
/// </summary>
public sealed class SpreadsheetPlugin(
    ILogger<SpreadsheetPlugin> logger,
    IFileStorageService fileStorage,
    IActivityPublisher activityPublisher,
    ISpreadsheetAnalysisService analysisService)
{
    private readonly ILogger<SpreadsheetPlugin> _logger =
        logger ?? throw new ArgumentNullException(nameof(logger));

    private readonly IFileStorageService _fileStorage =
        fileStorage ?? throw new ArgumentNullException(nameof(fileStorage));

    private readonly IActivityPublisher _activityPublisher =
        activityPublisher ?? throw new ArgumentNullException(nameof(activityPublisher));

    private readonly ISpreadsheetAnalysisService _analysisService =
        analysisService ?? throw new ArgumentNullException(nameof(analysisService));

    [KernelFunction("query_spreadsheet")]
    [Description("Queries Excel data using natural language with high accuracy through sandbox execution")]
    public async Task<string> QuerySpreadsheetAsync(
        [Description("Full path of the workbook")]
        string filePath,
        [Description("Natural language query")]
        string query,
        [Description("Sheet name (optional)")] string sheetName = "")
    {
        _logger.LogInformation("QuerySpreadsheet V3 Refactored: {File} - '{Query}'", filePath, query);

        // Log the start of the query process
        await _activityPublisher.PublishAsync("query_spreadsheet.start", new
        {
            filePath,
            query,
            sheetName,
            timestamp = DateTime.UtcNow,
            version = "V3Refactored"
        });

        try
        {
            // TODO: detect the sheet name if not provided - fallback to first sheet

            // Step 1: Load workbook
            using var workbook = await LoadWorkbookAsync(filePath);
            var sourceSheet = GetWorksheet(workbook, sheetName);

            // Step 2: Detect document format using stratified sampling
            var format = await _analysisService.DetectDocumentFormatAsync(sourceSheet);

            // Step 3: Validate format
            if (format != DocumentFormat.Columnar)
            {
                throw new InvalidOperationException(
                    $"Unsupported spreadsheet format: {format}. Only columnar format is currently supported.");
            }

            // Step 4: Extract headers from the source sheet
            var headers = _analysisService.ExtractHeaders(sourceSheet);
            await _activityPublisher.PublishAsync("query_spreadsheet.ExtractHeaders", new
            {
                filePath,
                sheetName = sourceSheet.Name,
                rowCount = sourceSheet.Cells.MaxRow + 1,
                columnCount = sourceSheet.Cells.MaxColumn + 1,
                headers
            });

            // Step 5: Extract metadata based on detected format
            var metadata = await _analysisService.ExtractDocumentMetadataAsync(sourceSheet, format, headers);
            
            // Log enhanced metadata information
            await _activityPublisher.PublishAsync("query_spreadsheet.metadata_extracted", new
            {
                filePath,
                sheetName = sourceSheet.Name,
                format,
                totalRows = metadata.TotalRows,
                dataRowCount = metadata.DataRowCount,
                headerRowIndex = metadata.DataStartRow - 1,
                dataStartRow = metadata.DataStartRow,
                dataEndRow = metadata.TotalRows - 1,
                totalColumns = metadata.TotalColumns,
                headers = metadata.Headers,
                columnStatistics = metadata.DataTypes.ContainsKey("_column_statistics") 
                    ? JsonSerializer.Deserialize<object>(metadata.DataTypes["_column_statistics"]) 
                    : null
            });

            // Step 6: Analyze the query and the navigate spreadsheet to get the necessary context
            var analysisResult = await _analysisService.AnalyzeQueryAsync(query, metadata, sourceSheet);

            // Extract execution plan from the artifact
            var executionPlan = ExtractExecutionPlan(analysisResult);

            // Enhanced logging with full ArtifactsFormatted data
            await _activityPublisher.PublishAsync("query_spreadsheet.AnalyzeQueryAsync", new
            {
                needRunFormula = executionPlan.NeedRunFormula,
                formula = executionPlan.Formula,
                simpleAnswer = executionPlan.SimpleAnswer,
                reasoning = executionPlan.Reasoning,
                artifactsRowCount = executionPlan.ArtifactsFormatted?.Count ?? 0,
                artifactFormatted = executionPlan.ArtifactsFormatted?.Select(row => string.Join(", ", row)).ToList() ??
                                    new List<string>(),
                // Add complete ArtifactsFormatted for debugging
                artifactsFormattedComplete = executionPlan.ArtifactsFormatted,
                artifactsFormattedJson = JsonSerializer.Serialize(executionPlan.ArtifactsFormatted, new JsonSerializerOptions { WriteIndented = true })
            });

            // Check if we need to run a formula or just return the simple answer
            if (!executionPlan.NeedRunFormula)
            {
                var simpleResult = new
                {
                    Success = true,
                    Query = query,
                    Answer = executionPlan.SimpleAnswer,
                    Reasoning = executionPlan.Reasoning,
                    RequiredCalculation = false,
                    DatasetContext = new
                    {
                        HeaderRowIndex = metadata.DataStartRow - 1,
                        DataStartRow = metadata.DataStartRow,
                        DataEndRow = metadata.TotalRows - 1,
                        TotalDataRows = metadata.DataRowCount,
                        Explanation =
                            $"The dataset has headers at row {metadata.DataStartRow - 1} and contains {metadata.DataRowCount} data rows (from row {metadata.DataStartRow} to {metadata.TotalRows - 1})"
                    }
                };

                await _activityPublisher.PublishAsync("query_spreadsheet.completed", new
                {
                    filePath,
                    query,
                    result = simpleResult,
                    success = true,
                    executionTime = DateTime.UtcNow,
                    simpleAnswer = true
                });

                PublishTool("query_spreadsheet", new { filePath, query, sheetName },
                    JsonSerializer.Serialize(simpleResult));

                return JsonSerializer.Serialize(simpleResult);
            }

            // Step 7: Create dynamic spreadsheet from artifacts
            var dynamicSpreadsheetResult = await _analysisService.CreateDynamicSpreadsheetAsync(executionPlan);
            var dynamicWorkbook = dynamicSpreadsheetResult.Workbook;
            var dynamicSheet = dynamicSpreadsheetResult.Worksheet;
            
            // Log all cell assignments
            await _activityPublisher.PublishAsync("query_spreadsheet.CellAssignments", new
            {
                totalCells = dynamicSpreadsheetResult.CellAssignments.Count,
                assignments = dynamicSpreadsheetResult.CellAssignments.Select(ca => new
                {
                    cellRef = ca.CellReference,
                    row = ca.Row,
                    col = ca.Column,
                    originalValue = ca.OriginalValue,
                    originalType = ca.OriginalType,
                    assignedValue = ca.AssignedValue,
                    assignedType = ca.AssignedType,
                    isJsonElement = ca.IsJsonElement
                })
            });

            // Enhanced logging for dynamic sheet creation with cell-by-cell data
            var cellData = new List<object>();
            if (executionPlan.ArtifactsFormatted != null)
            {
                for (int row = 0; row < executionPlan.ArtifactsFormatted.Count; row++)
                {
                    for (int col = 0; col < executionPlan.ArtifactsFormatted[row].Count; col++)
                    {
                        cellData.Add(new
                        {
                            row,
                            col,
                            value = executionPlan.ArtifactsFormatted[row][col],
                            type = executionPlan.ArtifactsFormatted[row][col]?.GetType().Name ?? "null",
                            cellReference = $"{(char)('A' + col)}{row + 1}"
                        });
                    }
                }
            }

            await _activityPublisher.PublishAsync("query_spreadsheet.DynamicSheetCreated", new
            {
                rows = dynamicSpreadsheetResult.DataRows,
                columns = dynamicSpreadsheetResult.DataColumns,
                formula = executionPlan.Formula,
                // Detailed cell data for debugging
                cellDetails = cellData,
                // Matrix view of the data
                dataMatrix = executionPlan.ArtifactsFormatted?.Select((row, rowIdx) => 
                    row.Select((cell, colIdx) => new 
                    { 
                        cell = $"{(char)('A' + colIdx)}{rowIdx + 1}", 
                        value = cell?.ToString() ?? "null",
                        type = cell?.GetType().Name ?? "null"
                    }).ToList()
                ).ToList()
            });

            try
            {
                // Step 8: Execute the formula
                var formulaExecution = await _analysisService.ExecuteFormulaAsync(dynamicWorkbook, executionPlan);
                
                // Log before formula execution
                await _activityPublisher.PublishAsync("query_spreadsheet.BeforeFormulaExecution", new
                {
                    formula = executionPlan.Formula,
                    formulaCellReference = formulaExecution.FormulaCellReference,
                    formulaRow = formulaExecution.DebugInfo["formulaRow"],
                    formulaColumn = formulaExecution.DebugInfo["formulaColumn"],
                    dataRowCount = formulaExecution.DebugInfo["dataRowCount"],
                    dataColumnCount = formulaExecution.DebugInfo["dataColumnCount"],
                    // Show the data range the formula will operate on
                    dataRange = executionPlan.ArtifactsFormatted != null && executionPlan.ArtifactsFormatted.Count > 0
                        ? $"A1:{(char)('A' + executionPlan.ArtifactsFormatted[0].Count - 1)}{executionPlan.ArtifactsFormatted.Count}"
                        : "Empty"
                });

                // Log after formula execution
                await _activityPublisher.PublishAsync("query_spreadsheet.AfterFormulaExecution", new
                {
                    formula = executionPlan.Formula,
                    formulaCellReference = formulaExecution.FormulaCellReference,
                    formulaValue = formulaExecution.Value,
                    formulaValueType = formulaExecution.Value?.GetType().Name ?? "null",
                    formulaResult = formulaExecution.StringValue,
                    formulaError = formulaExecution.Error,
                    // Show some sample values from the data used in calculation
                    sampleDataValues = executionPlan.ArtifactsFormatted?.Skip(1).Take(5)
                        .Select((row, idx) => new 
                        {
                            rowIndex = idx + 2, // +2 because we skip header and 0-based
                            values = row.Select(v => v?.ToString() ?? "null").ToList()
                        }).ToList()
                });

                var formulaResult = formulaExecution.StringValue;
                var formulaValue = formulaExecution.Value;

                // Step 9: Create final result
                var finalResult = new
                {
                    Success = true,
                    Query = query,
                    Answer = formulaResult,
                    Formula = executionPlan.Formula,
                    Reasoning = executionPlan.Reasoning,
                    RequiredCalculation = true,
                    DatasetContext = new
                    {
                        HeaderRowIndex = metadata.DataStartRow - 1,
                        DataStartRow = metadata.DataStartRow,
                        DataEndRow = metadata.TotalRows - 1,
                        TotalDataRows = metadata.DataRowCount,
                        Explanation =
                            $"The dataset has headers at row {metadata.DataStartRow - 1} and contains {metadata.DataRowCount} data rows (from row {metadata.DataStartRow} to {metadata.TotalRows - 1})"
                    },
                    DataUsed = new
                    {
                        Rows = executionPlan.ArtifactsFormatted?.Count ?? 0,
                        Columns = executionPlan.ArtifactsFormatted?.FirstOrDefault()?.Count ?? 0,
                        Headers = executionPlan.ArtifactsFormatted?.FirstOrDefault()
                    }
                };

                await _activityPublisher.PublishAsync("query_spreadsheet.completed", new
                {
                    filePath,
                    query,
                    result = finalResult,
                    success = true,
                    executionTime = DateTime.UtcNow,
                    formulaExecuted = true,
                    formulaResult
                });

                PublishTool("query_spreadsheet", new { filePath, query, sheetName },
                    JsonSerializer.Serialize(finalResult));

                return JsonSerializer.Serialize(finalResult);
            }
            finally
            {
                // Cleanup dynamic workbook
                dynamicWorkbook.Dispose();
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "QuerySpreadsheet V3 Refactored failed");

            await _activityPublisher.PublishAsync("query_spreadsheet.error", new
            {
                filePath,
                query,
                error = ex.Message,
                stackTrace = ex.StackTrace
            });

            var errorResult = new
            {
                Success = false,
                Error = ex.Message,
                Query = query
            };

            // Publish tool invocation even for errors
            PublishTool("query_spreadsheet", new { filePath, query, sheetName },
                JsonSerializer.Serialize(errorResult));

            return JsonSerializer.Serialize(errorResult);
        }
    }


    #region Private Helper Methods

    private async Task<Workbook> LoadWorkbookAsync(string path)
    {
        Stream stream;
        try
        {
            stream = await _fileStorage.GetFileAsync(Path.GetFileName(path));
        }
        catch (FileNotFoundException)
        {
            stream = File.OpenRead(path);
        }

        return new Workbook(stream);
    }

    private Worksheet GetWorksheet(Workbook workbook, string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return workbook.Worksheets[0];

        return workbook.Worksheets.FirstOrDefault(w =>
            w.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) ?? workbook.Worksheets[0];
    }

    private void PublishTool(string toolName, object parameters, string result)
    {
        try
        {
            _activityPublisher.PublishAsync("tool_invocation", new
            {
                tool = toolName,
                parameters,
                result
            }).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to publish tool invocation");
        }
    }

    private ExecutionPlanDto ExtractExecutionPlan(QueryAnalysisResult analysisResult)
    {
        using var doc = JsonDocument.Parse(analysisResult.Artifact);
        var executionPlanElement = doc.RootElement.GetProperty("ExecutionPlan");
        return JsonSerializer.Deserialize<ExecutionPlanDto>(executionPlanElement.GetRawText())
               ?? new ExecutionPlanDto();
    }

    #endregion
}