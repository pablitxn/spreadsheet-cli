using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using OpenAI.Chat;
using SpreadsheetCLI.Application.Interfaces;
using SpreadsheetCLI.Application.Interfaces.Spreadsheet;
using SpreadsheetCLI.Application.DTOs;
using SpreadsheetCLI.Domain.Entities;
using SpreadsheetCLI.Domain.Enums;
using SpreadsheetCLI.Domain.ValueObjects;
using SpreadsheetCLI.Infrastructure.Services;

namespace SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Services;

/// <summary>
/// Service for analyzing spreadsheet data with dynamic format recognition and iterative context gathering
/// </summary>
public class SpreadsheetAnalysisService(
    ILogger<SpreadsheetAnalysisService> logger,
    IChatCompletionService chatCompletion,
    IActivityPublisher activityPublisher,
    FileLoggerService fileLogger)
    : ISpreadsheetAnalysisService
{
    private readonly ILogger<SpreadsheetAnalysisService> _logger =
        logger ?? throw new ArgumentNullException(nameof(logger));

    private readonly IChatCompletionService _chatCompletion =
        chatCompletion ?? throw new ArgumentNullException(nameof(chatCompletion));

    private readonly IActivityPublisher _activityPublisher =
        activityPublisher ?? throw new ArgumentNullException(nameof(activityPublisher));

    private readonly FileLoggerService _fileLogger =
        fileLogger ?? throw new ArgumentNullException(nameof(fileLogger));

    private const int SampleSize = 50;
    private const int MaxIterations = 10;
    private const int MaxInitSampleColumns = 50;


    /// <summary>
    /// Context snapshot for each iteration of the analysis loop
    /// </summary>
    public class ContextSnapshot
    {
        public int RowIndex { get; set; }
        public int ColIndex { get; set; }
        public DocumentFormat FormatType { get; init; }
        public bool RequiredHeadersSatisfied { get; set; }
        public List<string> CollectedHeaders { get; init; } = [];
        public List<string> MissingHeaders { get; set; } = [];
        public Dictionary<string, CellStatsSummary> CellStats { get; init; } = new();
        public string ArtifactDigest { get; set; } = "";
        public double CoveragePercent { get; set; }
        public int IterationCount { get; set; }
    }

    /// <summary>
    /// Summary statistics for inspected cells
    /// </summary>
    public class CellStatsSummary
    {
        public List<string> SampleValues { get; set; } = [];
        public string InferredType { get; set; } = "";
        public string ValueRange { get; set; } = "";
        public int NonNullCount { get; set; }
    }


    /// <summary>
    /// Analyzing spreadsheet data with dynamic format recognition and iterative context gathering
    /// </summary>
    public async Task<QueryAnalysisResult> AnalyzeQueryAsync(
        string query,
        DocumentMetadata metadata,
        Worksheet worksheet,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Starting dynamic query analysis: {Query}", query);
        _logger.LogInformation(
            "Dataset structure - Headers at row: {HeaderRow}, Data from row {DataStart} to {DataEnd}, Total data rows: {DataRows}",
            metadata.DataStartRow - 1, metadata.DataStartRow, metadata.TotalRows - 1, metadata.DataRowCount);

        // Initialize variables for the analysis loop
        var artifacts = new List<string>();
        var currentRowIndex = 0;
        var iterationCount = 0;
        var continueAnalysis = true;
        var userIntentWithContext = "";

        // Use the data start row from metadata
        currentRowIndex = metadata.DataStartRow;

        // Iterative analysis loop
        while (continueAnalysis && iterationCount < MaxIterations && currentRowIndex <= worksheet.Cells.MaxRow)
        {
            iterationCount++;

            // Build markdown table with current batch of rows
            var markdownTable = BuildMarkdownTable(worksheet, metadata, currentRowIndex, SampleSize);

            // Analyze this batch with LLM
            var batchResult = await AnalyzeBatchWithLlmAsync(
                query,
                metadata,
                markdownTable,
                artifacts,
                iterationCount,
                cancellationToken);

            // Update artifacts with new findings
            if (!string.IsNullOrWhiteSpace(batchResult.NewArtifacts))
            {
                artifacts.Add(batchResult.NewArtifacts);
            }

            // Update user intent if provided
            if (!string.IsNullOrWhiteSpace(batchResult.UserIntentWithContext))
            {
                userIntentWithContext = batchResult.UserIntentWithContext;
            }

            // Check if we should continue
            continueAnalysis = batchResult.ContinueSnapshot;

            // Move to next batch
            currentRowIndex += SampleSize;

            await _activityPublisher.PublishAsync("spreadsheet_analysis.batch_complete", new
            {
                iteration = iterationCount,
                rowsProcessed = Math.Min(currentRowIndex, worksheet.Cells.MaxRow + 1),
                totalRows = worksheet.Cells.MaxRow + 1,
                dataRows = metadata.DataRowCount,
                artifactsCount = artifacts.Count,
                continueAnalysis,
                userIntentWithContext = batchResult.UserIntentWithContext
            });

            // Log batch analysis details for debugging
            await _fileLogger.LogDebugAsync("batch_analysis_result", new
            {
                iteration = iterationCount,
                batchResult,
                currentRowIndex,
                markdownTableLength = markdownTable.Length,
                artifactsCollected = artifacts
            });
        }

        // Combine all artifacts into a single artifact document
        var combinedArtifacts = string.Join("\n\n---\n\n", artifacts);

        // Generate final analysis result based on collected artifacts
        var finalResult = await GenerateExecutionPlanFromArtifactsAsync(
            query,
            userIntentWithContext,
            combinedArtifacts,
            metadata,
            cancellationToken);

        return finalResult;
    }

    /// <summary>
    /// Detects the document format using stratified sampling
    /// </summary>
    public async Task<DocumentFormat> DetectDocumentFormatAsync(
        Worksheet worksheet,
        CancellationToken cancellationToken)
    {
        var maxRow = worksheet.Cells.MaxRow;
        var maxCol = worksheet.Cells.MaxColumn;

        // Stratified sampling: first 50, middle 50, last 50 rows
        var samples = new List<string>();

        // First 50 rows
        for (int row = 0; row < Math.Min(SampleSize, maxRow + 1); row++)
        {
            var rowData = new List<string>();
            for (int col = 0; col <= Math.Min(10, maxCol); col++)
            {
                rowData.Add(worksheet.Cells[row, col].StringValue);
            }

            samples.Add(string.Join("|", rowData));
        }

        // Middle 50 rows
        if (maxRow > SampleSize * 2)
        {
            var middleStart = maxRow / 2 - SampleSize / 2;
            for (int row = middleStart; row < middleStart + SampleSize; row++)
            {
                var rowData = new List<string>();
                for (int col = 0; col <= Math.Min(10, maxCol); col++)
                {
                    rowData.Add(worksheet.Cells[row, col].StringValue);
                }

                samples.Add(string.Join("|", rowData));
            }
        }

        // Last 50 rows
        if (maxRow > SampleSize)
        {
            for (int row = Math.Max(0, maxRow - SampleSize + 1); row <= maxRow; row++)
            {
                var rowData = new List<string>();
                for (int col = 0; col <= Math.Min(10, maxCol); col++)
                {
                    rowData.Add(worksheet.Cells[row, col].StringValue);
                }

                samples.Add(string.Join("|", rowData));
            }
        }

        // Use LLM to detect format
        var prompt = $"""
                      Analyze these spreadsheet samples to determine the document format.

                      Samples (first 10 columns of selected rows):
                      {string.Join("\n", samples.Take(20))}

                      Determine if this is:
                      - Columnar: Traditional format with headers in first row
                      - RowBased: Data organized by rows with headers in first column
                      - Nested: Hierarchical or pivot-like structure
                      - Matrix: Cross-tabulation format
                      - Mixed: Combination of formats
                      - Unknown: Cannot determine format

                      Look for patterns like:
                      - Consistent headers in first row/column
                      - Repeating structures
                      - Hierarchical indentation
                      - Cross-references between rows and columns
                      """;

        var responseFormat = ChatResponseFormat.CreateJsonSchemaFormat(
            jsonSchemaFormatName: "format_detection",
            jsonSchema: BinaryData.FromString("""
                                              {
                                                  "type": "object",
                                                  "properties": {
                                                      "Format": { 
                                                          "type": "string",
                                                          "enum": ["Columnar", "RowBased", "Nested", "Matrix", "Mixed", "Unknown"]
                                                      },
                                                      "Confidence": { "type": "number" },
                                                      "Reasoning": { "type": "string" },
                                                      "HeaderLocation": { "type": "string" }
                                                  },
                                                  "required": ["Format", "Confidence", "Reasoning", "HeaderLocation"],
                                                  "additionalProperties": false
                                              }
                                              """),
            jsonSchemaIsStrict: true
        );

        var settings = new OpenAIPromptExecutionSettings
        {
            ModelId = "o4-mini",
            ResponseFormat = responseFormat,
            Temperature = 0.1
        };

        var chatHistory = new ChatHistory();
        chatHistory.AddMessage(AuthorRole.User, prompt);

        var response = await _chatCompletion.GetChatMessageContentsAsync(
            chatHistory, settings, cancellationToken: cancellationToken);

        var result = JsonSerializer.Deserialize<FormatDetectionResult>(response[0].Content ?? "{}");

        return Enum.Parse<DocumentFormat>(result?.Format ?? "Unknown");
    }

    /// <summary>
    /// Extracts document metadata based on the detected format
    /// </summary>
    public async Task<DocumentMetadata> ExtractDocumentMetadataAsync(
        Worksheet worksheet,
        DocumentFormat format,
        List<HeaderInfo> headers,
        CancellationToken cancellationToken)
    {
        // Get header row index from the first header (all headers should be on the same row)
        var headerRowIndex = headers.Any() ? headers.First().RowIndex : 0;
        var dataStartRow = headerRowIndex + 1;
        var totalRows = worksheet.Cells.MaxRow + 1;

        // Calculate data row count safely
        var dataRowCount = Math.Max(0, totalRows - dataStartRow);

        // Handle edge case where headers might be at the last row
        if (dataStartRow >= totalRows)
        {
            _logger.LogWarning(
                "Headers found at last row or beyond. No data rows available. HeaderRow: {HeaderRow}, TotalRows: {TotalRows}",
                headerRowIndex, totalRows);
            dataRowCount = 0;
        }

        var metadata = new DocumentMetadata
        {
            Format = format,
            TotalRows = totalRows,
            TotalColumns = worksheet.Cells.MaxColumn + 1,
            DataStartRow = dataStartRow,
            DataRowCount = dataRowCount
        };

        // Use the headers passed as parameter
        metadata.Headers = headers.Select(h => h.Name).ToList();

        // Detect data types and gather comprehensive statistics for each identified header/column
        metadata.DataTypes =
            await DetectDataTypesWithStatisticsAsync(worksheet, metadata, headers, format, cancellationToken);

        // Log enhanced metadata information
        await _activityPublisher.PublishAsync("metadata_extraction.enhanced", new
        {
            format,
            totalRows,
            dataRowCount,
            totalColumns = metadata.TotalColumns,
            headerCount = metadata.Headers.Count,
            dataTypes = metadata.DataTypes,
            timestamp = DateTime.UtcNow
        });

        return metadata;
    }

    /// <summary>
    /// Creates a markdown table representation of worksheet data with real row indices
    /// </summary>
    private string CreateMarkdownTableWithRealIndices(Worksheet worksheet, int startRow, int endRow,
        int? maxCols = null)
    {
        var sb = new StringBuilder();
        var actualMaxRow = worksheet.Cells.MaxRow;
        var actualMaxCol = worksheet.Cells.MaxColumn;

        // Validate and adjust row indices
        startRow = Math.Max(0, Math.Min(startRow, actualMaxRow));
        endRow = Math.Max(startRow, Math.Min(endRow, actualMaxRow));

        // Handle edge case where worksheet might be empty
        if (actualMaxRow < 0 || actualMaxCol < 0)
        {
            return "*Empty worksheet*";
        }

        var colsToShow = maxCols.HasValue ? Math.Min(maxCols.Value, actualMaxCol) : actualMaxCol;

        // Table header
        sb.Append("| Row # |");
        for (int col = 0; col <= colsToShow; col++)
        {
            sb.Append($" Col {col} |");
        }

        if (colsToShow < actualMaxCol)
        {
            sb.Append(" ... |");
        }

        sb.AppendLine();

        // Separator
        sb.Append("|-------|");
        for (int col = 0; col <= colsToShow; col++)
        {
            sb.Append("--------|");
        }

        if (colsToShow < actualMaxCol)
        {
            sb.Append("-----|");
        }

        sb.AppendLine();

        // Data rows with actual row indices
        var rowsShown = 0;
        for (int row = startRow; row <= endRow; row++)
        {
            sb.Append($"| {row} |");
            for (int col = 0; col <= colsToShow; col++)
            {
                var cellValue = "";
                try
                {
                    cellValue = worksheet.Cells[row, col].StringValue ?? "";
                }
                catch
                {
                    // Handle any access errors gracefully
                    cellValue = "";
                }

                // Limit cell content length and escape markdown
                cellValue = cellValue.Replace("|", "\\|").Replace("\n", " ").Trim();
                if (cellValue.Length > 30)
                {
                    cellValue = cellValue.Substring(0, 27) + "...";
                }

                sb.Append($" {cellValue} |");
            }

            if (colsToShow < actualMaxCol)
            {
                sb.Append(" ... |");
            }

            sb.AppendLine();
            rowsShown++;
        }

        // Add summary information
        if (rowsShown == 0)
        {
            sb.AppendLine("*No data rows in specified range*");
        }

        if (colsToShow < actualMaxCol)
        {
            sb.AppendLine($"*Note: Showing first {colsToShow + 1} columns of {actualMaxCol + 1} total columns*");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Uses LLM to extract headers from sample data
    /// </summary>
    private async Task<List<HeaderInfo>> ExtractHeadersWithLlmAsync(
        Worksheet worksheet,
        int maxSampleRows,
        CancellationToken cancellationToken)
    {
        var actualMaxRow = worksheet.Cells.MaxRow;
        var totalRows = actualMaxRow + 1;

        // Adjust sample size based on document size
        var rowsToShow = Math.Min(maxSampleRows - 1, actualMaxRow);

        // Create markdown table with real row indices
        var markdownTable = CreateMarkdownTableWithRealIndices(worksheet, 0, rowsToShow);

        var prompt = $"""
                      Analyze this spreadsheet data and identify the column headers.

                      IMPORTANT: The row numbers shown are the ACTUAL row indices from the Excel file.
                      Total rows in document: {totalRows}

                      {markdownTable}

                      Identify:
                      1. Which row contains the headers (look for the row with column names, not data)
                      2. Extract all column headers from that row
                      3. Return the ACTUAL row index where headers are found
                      4. If headers span multiple rows, combine them appropriately
                      5. Handle empty columns by using "Column_X" where X is the column index

                      Common patterns:
                      - Headers usually in row 0, but can be anywhere
                      - Look for rows with descriptive text that don't contain numeric data
                      - Headers might have empty rows above them

                      Example: If you see headers at row 18, return HeaderRowIndex: 18

                      For small documents (< 10 rows), be extra careful to distinguish headers from data.
                      """;

        var responseFormat = ChatResponseFormat.CreateJsonSchemaFormat(
            jsonSchemaFormatName: "header_extraction",
            jsonSchema: BinaryData.FromString("""
                                              {
                                                  "type": "object",
                                                  "properties": {
                                                      "HeaderRowIndex": { "type": "integer" },
                                                      "Headers": {
                                                          "type": "array",
                                                          "items": { "type": "string" }
                                                      },
                                                      "MultiRowHeaders": { "type": "boolean" },
                                                      "Confidence": { "type": "number" }
                                                  },
                                                  "required": ["HeaderRowIndex", "Headers", "MultiRowHeaders", "Confidence"],
                                                  "additionalProperties": false
                                              }
                                              """),
            jsonSchemaIsStrict: true
        );

        var settings = new OpenAIPromptExecutionSettings
        {
            ModelId = "gpt-4o-mini",
            ResponseFormat = responseFormat,
            Temperature = 0.1
        };

        var chatHistory = new ChatHistory();
        chatHistory.AddMessage(AuthorRole.User, prompt);

        var response = await _chatCompletion.GetChatMessageContentsAsync(
            chatHistory, settings, cancellationToken: cancellationToken);

        var result = JsonSerializer.Deserialize<HeaderExtractionResult>(response[0].Content ?? "{}");

        // Convert headers to HeaderInfo list with row index
        var headerInfoList = new List<HeaderInfo>();
        if (result != null && result.Headers != null)
        {
            for (int i = 0; i < result.Headers.Count; i++)
            {
                headerInfoList.Add(new HeaderInfo(result.Headers[i], result.HeaderRowIndex));
            }
        }

        return headerInfoList;
    }


    /// <summary>
    /// Detects data types and gathers comprehensive statistics for columns
    /// </summary>
    private Task<Dictionary<string, string>> DetectDataTypesWithStatisticsAsync(
        Worksheet worksheet,
        DocumentMetadata metadata,
        List<HeaderInfo> headers,
        DocumentFormat format,
        CancellationToken cancellationToken)
    {
        var dataTypes = new Dictionary<string, string>();
        var columnStatistics = new Dictionary<string, object>();

        foreach (var header in headers)
        {
            var colIndex = headers.IndexOf(header);
            var columnValues = new List<object>();
            var numericValues = new List<double>();
            var dateValues = new List<DateTime>();
            var textValues = new List<string>();
            var nonNullCount = 0;
            var uniqueValues = new HashSet<object>();

            // Collect all values in the column
            if (format == DocumentFormat.Columnar)
            {
                for (int row = metadata.DataStartRow; row < metadata.TotalRows && row <= worksheet.Cells.MaxRow; row++)
                {
                    var value = worksheet.Cells[row, colIndex].Value;
                    if (value != null)
                    {
                        columnValues.Add(value);
                        nonNullCount++;
                        uniqueValues.Add(value);

                        // Try to parse as numeric
                        if (TryParseNumeric(value, out var numValue))
                        {
                            numericValues.Add(numValue);
                        }
                        // Try to parse as date
                        else if (TryParseDate(value) is DateTime dateValue)
                        {
                            dateValues.Add(dateValue);
                        }
                        // Store as text
                        else
                        {
                            textValues.Add(value.ToString() ?? "");
                        }
                    }
                }
            }

            // Determine primary data type
            var inferredType =
                InferDataTypeFromDistribution(numericValues.Count, dateValues.Count, textValues.Count, nonNullCount);
            dataTypes[header.Name] = inferredType;

            // Calculate statistics based on data type
            var stats = new Dictionary<string, object>
            {
                ["NonNullCount"] = nonNullCount,
                ["NullCount"] = metadata.DataRowCount - nonNullCount,
                ["UniqueCount"] = uniqueValues.Count,
                ["DataType"] = inferredType,
                ["FillRate"] = metadata.DataRowCount > 0 ? (double)nonNullCount / metadata.DataRowCount : 0
            };

            if (inferredType == "numeric" && numericValues.Any())
            {
                // Calculate numeric statistics
                stats["Min"] = numericValues.Min();
                stats["Max"] = numericValues.Max();
                stats["Mean"] = numericValues.Average();
                stats["Sum"] = numericValues.Sum();

                // Calculate standard deviation
                var mean = numericValues.Average();
                var sumOfSquares = numericValues.Sum(v => Math.Pow(v - mean, 2));
                var stdDev = Math.Sqrt(sumOfSquares / numericValues.Count);
                stats["StdDev"] = stdDev;

                // Calculate variance
                stats["Variance"] = Math.Pow(stdDev, 2);

                // Calculate median
                var sortedValues = numericValues.OrderBy(v => v).ToList();
                stats["Median"] = sortedValues.Count % 2 == 0
                    ? (sortedValues[sortedValues.Count / 2 - 1] + sortedValues[sortedValues.Count / 2]) / 2
                    : sortedValues[sortedValues.Count / 2];

                // Calculate mode (most frequent value)
                var valueCounts = numericValues.GroupBy(v => v).OrderByDescending(g => g.Count()).FirstOrDefault();
                if (valueCounts != null && valueCounts.Count() > 1)
                {
                    stats["Mode"] = valueCounts.Key;
                    stats["ModeCount"] = valueCounts.Count();
                }

                // Calculate percentiles
                stats["Percentile25"] = CalculatePercentile(sortedValues, 0.25);
                stats["Percentile75"] = CalculatePercentile(sortedValues, 0.75);
                stats["IQR"] = (double)stats["Percentile75"] - (double)stats["Percentile25"];

                // Detect outliers using IQR method
                var iqr = (double)stats["IQR"];
                var q1 = (double)stats["Percentile25"];
                var q3 = (double)stats["Percentile75"];
                var lowerBound = q1 - 1.5 * iqr;
                var upperBound = q3 + 1.5 * iqr;
                var outliers = numericValues.Where(v => v < lowerBound || v > upperBound).ToList();
                stats["OutlierCount"] = outliers.Count;
                stats["OutlierPercentage"] =
                    numericValues.Count > 0 ? (double)outliers.Count / numericValues.Count * 100 : 0;

                // Sample values for context
                stats["SampleValues"] = numericValues.Take(10).ToList();
            }
            else if (inferredType == "date" && dateValues.Any())
            {
                // Calculate date statistics
                stats["MinDate"] = dateValues.Min();
                stats["MaxDate"] = dateValues.Max();
                stats["DateRange"] = (dateValues.Max() - dateValues.Min()).TotalDays;
                stats["SampleValues"] = dateValues.Take(10).Select(d => d.ToString("yyyy-MM-dd")).ToList();
            }
            else if (inferredType == "text" && textValues.Any())
            {
                // Calculate text statistics
                var lengths = textValues.Select(t => t.Length).ToList();
                stats["MinLength"] = lengths.Min();
                stats["MaxLength"] = lengths.Max();
                stats["AvgLength"] = lengths.Average();
                stats["SampleValues"] = textValues.Take(10).ToList();

                // Most common values
                var topValues = textValues.GroupBy(v => v)
                    .OrderByDescending(g => g.Count())
                    .Take(5)
                    .Select(g => new { Value = g.Key, Count = g.Count() })
                    .ToList();
                stats["TopValues"] = topValues;
            }

            // Store enhanced statistics in metadata
            if (!metadata.DataTypes.ContainsKey($"{header.Name}_stats"))
            {
                metadata.DataTypes[$"{header.Name}_stats"] = JsonSerializer.Serialize(stats);
            }

            columnStatistics[header.Name] = stats;
        }

        // Store overall column statistics in metadata
        metadata.DataTypes["_column_statistics"] = JsonSerializer.Serialize(columnStatistics);

        return Task.FromResult(dataTypes);
    }

    /// <summary>
    /// Calculates percentile from sorted list
    /// </summary>
    private double CalculatePercentile(List<double> sortedValues, double percentile)
    {
        if (!sortedValues.Any()) return 0;

        var index = percentile * (sortedValues.Count - 1);
        var lower = Math.Floor(index);
        var upper = Math.Ceiling(index);

        if (lower == upper)
        {
            return sortedValues[(int)index];
        }

        var weight = index - lower;
        return sortedValues[(int)lower] * (1 - weight) + sortedValues[(int)upper] * weight;
    }

    /// <summary>
    /// Infers data type based on value distribution
    /// </summary>
    private string InferDataTypeFromDistribution(int numericCount, int dateCount, int textCount, int totalCount)
    {
        if (totalCount == 0) return "unknown";

        var numericRatio = (double)numericCount / totalCount;
        var dateRatio = (double)dateCount / totalCount;
        var textRatio = (double)textCount / totalCount;

        // If more than 80% of values are numeric, consider it numeric
        if (numericRatio > 0.8) return "numeric";

        // If more than 80% of values are dates, consider it date
        if (dateRatio > 0.8) return "date";

        // If more than 50% are text or mixed types
        if (textRatio > 0.5) return "text";

        // Mixed type if no clear majority
        return "mixed";
    }

    #region Legacy Methods (Maintained for compatibility)

    /// <inheritdoc/>
    public List<HeaderInfo> ExtractHeaders(Worksheet worksheet)
    {
        return Task.Run(async () => await ExtractHeadersAsync(worksheet)).GetAwaiter().GetResult();
    }

    /// <summary>
    /// Extrae headers usando las primeras 50 filas y análisis con LLM
    /// </summary>
    private async Task<List<HeaderInfo>> ExtractHeadersAsync(Worksheet worksheet)
    {
        var maxRows = Math.Min(MaxInitSampleColumns, worksheet.Cells.MaxRow + 1);
        var headers = await ExtractHeadersWithLlmAsync(worksheet, maxRows, CancellationToken.None);
        return headers;
    }

    /// <inheritdoc/>
    public int CountDataRows(Worksheet worksheet)
    {
        var range = GetDataRange(worksheet);
        return range.LastRow - range.FirstRow + 1;
    }

    /// <inheritdoc/>
    public (int FirstRow, int LastRow) GetDataRange(Worksheet worksheet)
    {
        return (1, worksheet.Cells.MaxRow);
    }

    /// <inheritdoc/>
    public bool RowMatchesFilters(
        Worksheet worksheet,
        int row,
        List<HeaderInfo> headers,
        List<FilterCriteria> filters)
    {
        if (!filters.Any()) return true;

        foreach (var filter in filters)
        {
            var colIndex = headers.FindIndex(h => h.Name.Equals(filter.Column, StringComparison.OrdinalIgnoreCase));
            if (colIndex < 0) continue;

            var cellValue = worksheet.Cells[row, colIndex].Value;
            bool matches = EvaluateFilter(cellValue, filter);

            if (!matches) return false;
        }

        return true;
    }

    #endregion

    /// <summary>
    /// Builds a markdown table from spreadsheet data including row and column indices
    /// </summary>
    private string BuildMarkdownTable(Worksheet worksheet, DocumentMetadata metadata, int startRow, int rowCount)
    {
        var sb = new StringBuilder();
        var actualMaxRow = worksheet.Cells.MaxRow;
        var actualMaxCol = worksheet.Cells.MaxColumn;

        // Validate start row - FIX: No modificar startRow si es válido
        if (startRow < 0) startRow = 0;
        if (startRow > actualMaxRow)
        {
            sb.AppendLine("### No more data to analyze");
            sb.AppendLine($"Requested start row {startRow} is beyond the last row {actualMaxRow}");
            return sb.ToString();
        }

        var endRow = Math.Min(startRow + rowCount - 1, actualMaxRow);

        // Handle case where there's no data to show
        if (rowCount <= 0)
        {
            sb.AppendLine("### No data requested");
            return sb.ToString();
        }

        // Limit columns for readability
        var colsToShow = Math.Min(actualMaxCol, 20);

        // Table header with column letters
        sb.AppendLine("### Data Sample");
        sb.AppendLine($"*Showing rows {startRow} to {endRow}*");
        sb.AppendLine();

        sb.Append("| Row | ");
        for (int col = 0; col <= colsToShow; col++)
        {
            var colLetter = GetColumnLetter(col);
            sb.Append($"{colLetter}");
            if (col < metadata.Headers.Count && !string.IsNullOrEmpty(metadata.Headers[col]))
            {
                var header = metadata.Headers[col].Trim();
                if (header.Length > 15) header = header.Substring(0, 12) + "...";
                sb.Append($"<br/>*{header}*");
            }

            sb.Append(" | ");
        }

        if (colsToShow < actualMaxCol)
        {
            sb.Append("... | ");
        }

        sb.AppendLine();

        // Table separator
        sb.Append("|:---:|");
        for (int col = 0; col <= colsToShow; col++)
        {
            sb.Append(":------:|");
        }

        if (colsToShow < actualMaxCol)
        {
            sb.Append(":---:|");
        }

        sb.AppendLine();

        var rowsShown = 0;
        for (int row = startRow; row <= endRow && row <= actualMaxRow; row++)
        {
            sb.Append($"| **{row}** | ");

            for (int col = 0; col <= colsToShow; col++)
            {
                var cellValue = "";
                try
                {
                    var cell = worksheet.Cells[row, col];
                    if (cell != null)
                    {
                        cellValue = cell.StringValue ?? cell.Value?.ToString() ?? "";
                    }
                }
                catch
                {
                    cellValue = "";
                }

                // Escape markdown special characters and limit cell length
                cellValue = cellValue.Replace("|", "\\|")
                    .Replace("\n", " ")
                    .Replace("\r", "")
                    .Trim();

                if (cellValue.Length > 50)
                {
                    cellValue = cellValue.Substring(0, 47) + "...";
                }

                // Empty cells show as light gray
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    cellValue = "_empty_";
                }

                sb.Append($"{cellValue} | ");
            }

            if (colsToShow < actualMaxCol)
            {
                sb.Append("... | ");
            }

            sb.AppendLine();
            rowsShown++;
        }

        sb.AppendLine();
        sb.AppendLine("### 📊 Dataset Info");
        sb.AppendLine();
        sb.AppendLine("**Structure:**");
        sb.AppendLine($"- Headers at row: **{metadata.DataStartRow - 1}**");
        sb.AppendLine($"- Data range: rows **{metadata.DataStartRow}** to **{metadata.TotalRows - 1}**");
        sb.AppendLine($"- Total data rows: **{metadata.DataRowCount:N0}** *(excluding headers)*");
        sb.AppendLine(
            $"- Total columns: **{actualMaxCol + 1}** ({GetColumnLetter(0)} to {GetColumnLetter(actualMaxCol)})");

        sb.AppendLine();
        sb.AppendLine("**Current View:**");
        sb.AppendLine($"- Showing: rows **{startRow}** to **{endRow}** ({rowsShown} rows)");

        if (colsToShow < actualMaxCol)
        {
            sb.AppendLine($"- Columns: **{colsToShow + 1}** of **{actualMaxCol + 1}** (truncated for readability)");
        }

        // Progress indicator
        var progress = ((double)(endRow - metadata.DataStartRow + 1) / metadata.DataRowCount) * 100;
        sb.AppendLine($"- Progress: **{progress:F1}%** of data analyzed");

        return sb.ToString();
    }

// Helper method to convert column index to Excel-style letter
    private string GetColumnLetter(int columnIndex)
    {
        var dividend = columnIndex + 1;
        var columnName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

    /// <summary>
    /// Analyzes a batch of data using LLM
    /// </summary>
    private async Task<BatchAnalysisResult> AnalyzeBatchWithLlmAsync(
        string query,
        DocumentMetadata metadata,
        string markdownTable,
        List<string> previousArtifacts,
        int iteration,
        CancellationToken cancellationToken)
    {
        var previousArtifactsText = previousArtifacts.Any()
            ? $"\n\nSome previous artifacts collected:\n{string.Join("\n---\n", previousArtifacts.Take(3))}"
            : "";

        var prompt = $"""
                      You are analyzing spreadsheet data to answer this query: {query}

                      CRITICAL STRUCTURAL CONTEXT:
                      - Headers are located at row: {metadata.DataStartRow - 1}
                      - Data starts at row: {metadata.DataStartRow}
                      - Data ends at row: {metadata.TotalRows - 1}
                      - Total DATA rows (excluding headers): {metadata.DataRowCount}
                      - This means: Row indices {metadata.DataStartRow} through {metadata.TotalRows - 1} contain the actual data

                      Document metadata:
                      - Format: {metadata.Format}
                      - Total rows in file: {metadata.TotalRows}
                      - Total columns: {metadata.TotalColumns}
                      - Headers: {string.Join(", ", metadata.Headers)}
                      - Data types: {JsonSerializer.Serialize(metadata.DataTypes.Where(kvp => !kvp.Key.EndsWith("_stats") && kvp.Key != "_column_statistics").ToDictionary(kvp => kvp.Key, kvp => kvp.Value))}

                      Column Statistics Available:
                      {(metadata.DataTypes.ContainsKey("_column_statistics") ? metadata.DataTypes["_column_statistics"] : "No statistics available")}

                      This is iteration {iteration} of the analysis.{previousArtifactsText}

                      Current data sample:
                      {markdownTable}

                      IMPORTANT FOR CALCULATIONS:
                      - The dataset has {metadata.DataRowCount} data rows (this is what matters for percentages!)
                      - Row indices shown above are the actual Excel row numbers
                      - When calculating percentages: denominator = {metadata.DataRowCount} (NOT {metadata.TotalRows})

                      FOR PERCENTAGE QUERIES - You MUST:
                      1. Extract ALL rows that match the condition from the current batch
                      2. Keep a running count of matching rows
                      3. Remember that percentage = (matching rows / {metadata.DataRowCount}) * 100
                      4. Include these critical numbers in your UserIntentWithContext

                      Instructions:
                      1. Extract any data from this sample that is relevant to answering the user's query
                      2. Store the extracted data as "artifacts" - these should be small tables or data snippets in plain text
                      3. If the query requires multiple calculations (e.g., "average of X compared to sum of Y"), create separate artifact sections for each
                      4. Determine if you need to continue analyzing more rows or if you have sufficient data
                      5. Provide the user intent with COMPLETE structural context
                      6. For queries about percentages or proportions, ensure you collect enough data to make accurate calculations
                      
                      CRITICAL FOR AGGREGATE FUNCTIONS (MIN, MAX, AVERAGE, PERCENTILE, VARIANCE, etc.):
                      - If the query involves these functions on a column, you MUST extract ALL values from that column
                      - Continue analysis until you have collected the ENTIRE column data
                      - Do not stop at "first evidence" - range functions need the complete dataset
                      - For MIN/MAX of dates: Extract ALL date values, not just a sample

                      Format artifacts like this:
                      ## Artifact: [Description]
                      [Data in a simple format, could be a small table or list]

                      CRITICAL for UserIntentWithContext:
                      Include ALL structural observations:
                      - "The user wants to know what percentage of rows have Quantity > 1000"
                      - "The dataset has headers at row {metadata.DataStartRow - 1} and data from rows {metadata.DataStartRow} to {metadata.TotalRows - 1}"
                      - "This gives us {metadata.DataRowCount} total data rows for percentage calculations"
                      - "I found X rows matching the condition out of {metadata.DataRowCount} total data rows"

                      Example artifact format for percentage queries:
                      ## Artifact: Rows with Quantity > 1000
                      Row | Quantity
                      {metadata.DataStartRow + 3} | 1200
                      {metadata.DataStartRow + 6} | 1500
                      ...
                      Total matching rows found: [count]
                      Total data rows in dataset: {metadata.DataRowCount}
                      Header row index: {metadata.DataStartRow - 1}
                      Data row range: {metadata.DataStartRow} to {metadata.TotalRows - 1}
                      """;

        var responseFormat = ChatResponseFormat.CreateJsonSchemaFormat(
            jsonSchemaFormatName: "batch_analysis",
            jsonSchema: BinaryData.FromString("""
                                              {
                                                  "type": "object",
                                                  "properties": {
                                                      "NewArtifacts": {
                                                          "type": "string",
                                                          "description": "New artifacts extracted from this batch"
                                                      },
                                                      "ContinueSnapshot": {
                                                          "type": "boolean",
                                                          "description": "Whether to continue analyzing more rows"
                                                      },
                                                      "UserIntentWithContext": {
                                                          "type": "string",
                                                          "description": "The user's intent with full context from data seen so far"
                                                      },
                                                      "Reasoning": {
                                                          "type": "string",
                                                          "description": "Explanation of what was found and why to continue or stop"
                                                      }
                                                  },
                                                  "required": ["NewArtifacts", "ContinueSnapshot", "UserIntentWithContext", "Reasoning"],
                                                  "additionalProperties": false
                                              }
                                              """),
            jsonSchemaIsStrict: true
        );

        var settings = new OpenAIPromptExecutionSettings
        {
            ModelId = "o4-mini",
            ResponseFormat = responseFormat,
            Temperature = 0.2
        };

        var chatHistory = new ChatHistory();
        chatHistory.AddMessage(AuthorRole.User, prompt);

        var response = await _chatCompletion.GetChatMessageContentsAsync(
            chatHistory, settings, cancellationToken: cancellationToken);

        var result = JsonSerializer.Deserialize<BatchAnalysisResult>(response[0].Content ?? "{}");

        // Log batch LLM analysis details
        await _fileLogger.LogDebugAsync("batch_llm_analysis", new
        {
            iteration,
            promptLength = prompt.Length,
            modelResponse = response[0].Content,
            parsedResult = result,
            continueSnapshot = result?.ContinueSnapshot,
            newArtifactsLength = result?.NewArtifacts?.Length ?? 0
        });

        return result ?? new BatchAnalysisResult();
    }

    /// <summary>
    /// Generates execution plan from collected artifacts
    /// </summary>
    private async Task<QueryAnalysisResult> GenerateExecutionPlanFromArtifactsAsync(
        string query,
        string userIntentWithContext,
        string combinedArtifacts,
        DocumentMetadata metadata,
        CancellationToken cancellationToken)
    {
        var prompt = $"""
                      Generate an execution plan for this query: {query}

                      User intent with FULL STRUCTURAL CONTEXT: {userIntentWithContext}

                      Use Excel standard functions in English following the ISO/IEC 29500 specification, 
                      ensuring correct syntax, comma-separated arguments, and formulas always beginning with "="; 
                      verify compatibility with Aspose.Cells.

                      CRITICAL: Formula Generation for Dynamic Spreadsheet
                      - You will be creating a formula that operates on a NEW spreadsheet created from ArtifactsFormatted
                      - The dynamic spreadsheet contains ONLY the data you put in ArtifactsFormatted (not the full original dataset!)
                      - The dynamic spreadsheet will have columns starting from A, B, C, etc.
                      - DO NOT use column references from the original file 
                      - Map your columns based on the order in ArtifactsFormatted:
                        * First column in ArtifactsFormatted = Column A
                        * Second column in ArtifactsFormatted = Column B
                        * Third column in ArtifactsFormatted = Column C, etc.
                      - Data starts at row 2 (row 1 contains headers)
                      - IMPORTANT: If you extracted 100 matching rows, the dynamic sheet has 101 rows (1 header + 100 data)
                      
                      FOR RANGE/AGGREGATE FUNCTIONS:
                      - When using MIN, MAX, AVERAGE, VAR.S, VAR.P, PERCENTILE, etc., ensure ArtifactsFormatted contains ALL relevant data
                      - For date columns: Include ALL date values to avoid MIN returning 0
                      - Use the full range: e.g., =MIN(A2:A23) if you have 22 data rows

                      Always for any mathematical calculations return `NeedRunFormula=true`:


                      CRITICAL UNDERSTANDING:
                      - The UserIntentWithContext above contains crucial structural observations from the iterative analysis
                      - It should tell you exactly how many matching rows were found and what the total data row count is
                      - The artifacts contain the actual matching data extracted during traversal
                      - For percentage calculations: Use the counts mentioned in UserIntentWithContext!

                      VARIANCE AND GROUP CALCULATIONS:
                      - For "highest variance per group": Calculate variance for each group, identify the highest
                      - Use VAR.S for sample variance (n-1 denominator) unless specifically asked for population variance
                      - Clean numeric data before variance: Round quantities to reasonable precision
                      - IMPORTANT: Return only the numeric variance value in SimpleAnswer, not a narrative sentence
                      - In Reasoning, you can explain which group has the highest variance
                      
                      OUTPUT FORMATTING:
                      - Round results to 2 decimal places for display
                      - Format large numbers with thousand separators
                      - For variance values, use scientific notation if > 1e9
                      
                      ANSWER SEPARATION (CRITICAL):
                      - MachineAnswer: ONLY the numeric result or direct answer (e.g., "291839462.7", "2021-11-15", "MSFT")
                      - HumanExplanation: Full narrative with context (e.g., "SecurityGroup 1 has the highest variance of 291839462.7")
                      - For group-based queries asking "which X has highest/lowest Y": 
                        * MachineAnswer = just the Y value
                        * HumanExplanation = "X has Y value of [amount]"
                      
                      Instructions:
                      1. If the query can be answered directly from the artifacts (zero math needed), provide a simple_answer
                      2. For group-based queries, always include both the group identifier and the calculated value
                      3. Always provide reasoning that references the structural context

                      For ArtifactsFormatted, structure the data as a 2D array where:
                      - First row contains headers
                      - Subsequent rows contain the data values
                      - Include only the columns needed for the calculation
                      - Ensure data types are preserved (numbers as numbers, not strings)
                      
                      DATA CLEANSING NOTES:
                      - Convert accounting negatives "(123)" or "123-" to proper negative numbers: -123
                      - When asked for "negative values", interpret accounting formats correctly
                      - Clean currency symbols and thousand separators from numeric values

                      FORMULA GUIDELINES based on UserIntentWithContext:
                      - Read the UserIntentWithContext carefully - it contains the exact counts!
                      - For percentage queries where UserIntentWithContext says "I found X rows matching out of Y total data rows":
                        * Simple answer approach: Just calculate X/Y*100 directly and put in SimpleAnswer
                        * Formula approach: Since the dynamic spreadsheet only contains matching rows, use:
                          - If calculating percentage: =(COUNTA(data_range)-1)/{metadata.DataRowCount}*100
                          - Note: -1 to exclude header row, and {metadata.DataRowCount} is the total from original dataset
                        * IMPORTANT: For SUM/AVERAGE queries on the dynamic spreadsheet, formulas operate only on the extracted data
                      - The key insight: UserIntentWithContext already did the hard work of counting!
                      
                      RATIO CALCULATIONS:
                      - Row-wise ratio: Average of individual row ratios = AVERAGE(C2:C17) where C contains =A/B for each row
                      - Aggregate ratio (weighted): Total sum divided = SUM(Income)/SUM(Quantity)
                      - DEFAULT: Use aggregate ratio unless explicitly asked for "average ratio per row"

                      Example reasoning for percentage query:
                      "Based on the UserIntentWithContext, we found X rows with Quantity > 1000 out of {metadata.DataRowCount} total data rows.
                      The percentage is therefore (X / {metadata.DataRowCount}) * 100 = Z%"

                      NEVER confuse:
                      - Total rows in file ({metadata.TotalRows}) with total data rows ({metadata.DataRowCount})
                      - Row indices with row counts
                      - The header row is at index {metadata.DataStartRow - 1}, but this doesn't affect our percentage calculation


                      Collected artifacts from document traversal:
                      {combinedArtifacts}

                      Document metadata:
                      - Headers: {string.Join(", ", metadata.Headers)}
                      - Data types: {JsonSerializer.Serialize(metadata.DataTypes.Where(kvp => !kvp.Key.EndsWith("_stats") && kvp.Key != "_column_statistics").ToDictionary(kvp => kvp.Key, kvp => kvp.Value))}
                      - Total rows in file: {metadata.TotalRows}

                      Column Statistics for Calculations:
                      {(metadata.DataTypes.ContainsKey("_column_statistics") ? metadata.DataTypes["_column_statistics"] : "No statistics available")}

                      """;

        var responseFormat = ChatResponseFormat.CreateJsonSchemaFormat(
            jsonSchemaFormatName: "execution_plan",
            jsonSchema: BinaryData.FromString("""
                                              {
                                                "type": "object",
                                                "properties": {
                                                  "NeedRunFormula": {
                                                    "type": "boolean",
                                                    "description": "Always for ANY mathematical calculations return True"
                                                  },
                                                  "ArtifactsFormatted": {
                                                    "type": "array",
                                                    "description": "2D array representing the Excel data. First row is headers.",
                                                    "items": {
                                                      "type": "array",
                                                      "items": {
                                                        "type": ["string", "number", "boolean", "null"]
                                                      }
                                                    }
                                                  },
                                                  "Formula": {
                                                    "type": "string",
                                                    "description": "Formula Generation for Dynamic Spreadsheet"
                                                  },
                                                  "SimpleAnswer": {
                                                    "type": "string",
                                                    "description": "Direct answer if the query can be answered"
                                                  },
                                                  "Reasoning": {
                                                    "type": "string",
                                                    "description": "Explanation of the approach taken"
                                                  },
                                                  "MachineAnswer": {
                                                    "type": "string",
                                                    "description": "Machine-readable answer: just the numeric value/result, no narrative"
                                                  },
                                                  "HumanExplanation": {
                                                    "type": "string",
                                                    "description": "Human-readable explanation with context (e.g., 'Group X has variance Y')"
                                                  }
                                                },
                                                "required": [
                                                  "NeedRunFormula",
                                                  "ArtifactsFormatted",
                                                  "Formula",
                                                  "SimpleAnswer",
                                                  "Reasoning",
                                                  "MachineAnswer",
                                                  "HumanExplanation"
                                                ],
                                                "additionalProperties": false
                                              }
                                              """),
            jsonSchemaIsStrict: true
        );


        var settings = new OpenAIPromptExecutionSettings
        {
            ModelId = "o4-mini-high",
            ResponseFormat = responseFormat,
            Temperature = 0.1
        };

        var chatHistory = new ChatHistory();
        chatHistory.AddMessage(AuthorRole.User, prompt);

        var response = await _chatCompletion.GetChatMessageContentsAsync(
            chatHistory, settings, cancellationToken: cancellationToken);

        var result = JsonSerializer.Deserialize<ExecutionPlanResponse>(response[0].Content ?? "{}");

        // Log execution plan generation details
        await _fileLogger.LogDebugAsync("execution_plan_generated", new
        {
            query,
            userIntentWithContext,
            combinedArtifactsLength = combinedArtifacts.Length,
            modelResponse = response[0].Content,
            parsedResult = result,
            artifactsFormatted = result?.ArtifactsFormatted,
            formula = result?.Formula,
            simpleAnswer = result?.SimpleAnswer,
            reasoning = result?.Reasoning
        });

        // Convert the execution plan to QueryAnalysisResult
        // This is a temporary mapping until we update the return type
        return new QueryAnalysisResult
        {
            ColumnsNeeded = ExtractColumnsFromArtifacts(result?.ArtifactsFormatted),
            Filters = new List<FilterCriteria>(), // No filters in the new approach
            AggregationType = "", // Determined by formula
            GroupBy = null,
            RequiresCalculation = result?.NeedRunFormula ?? false,
            CalculationSteps = new List<string> { result?.Formula ?? "" },
            RequiresFullDataset = false,
            UserIntentWithContext = userIntentWithContext,
            Artifact = JsonSerializer.Serialize(new
            {
                ExecutionPlan = result,
                OriginalArtifacts = combinedArtifacts
            }),
        };
    }

    /// <summary>
    /// Extracts column names from formatted artifacts
    /// </summary>
    private List<string> ExtractColumnsFromArtifacts(List<List<object>>? artifacts)
    {
        if (artifacts == null || artifacts.Count == 0)
            return new List<string>();

        // First row contains headers
        return artifacts[0]?.Select(h => h?.ToString() ?? "").ToList() ?? new List<string>();
    }

    /// <summary>
    /// Creates a dynamic spreadsheet from execution plan data
    /// </summary>
    public async Task<DynamicSpreadsheetResult> CreateDynamicSpreadsheetAsync(
        ExecutionPlanDto executionPlan,
        CancellationToken cancellationToken = default)
    {
        var dynamicWorkbook = new Workbook();
        var dynamicSheet = dynamicWorkbook.Worksheets[0];
        dynamicSheet.Name = "DynamicData";

        var cellAssignments = new List<CellAssignment>();
        int dataRows = 0;
        int dataColumns = 0;

        if (executionPlan.ArtifactsFormatted != null && executionPlan.ArtifactsFormatted.Count > 0)
        {
            dataRows = executionPlan.ArtifactsFormatted.Count;
            dataColumns = executionPlan.ArtifactsFormatted[0].Count;

            for (int row = 0; row < dataRows; row++)
            {
                var rowData = executionPlan.ArtifactsFormatted[row];
                for (int col = 0; col < rowData.Count; col++)
                {
                    var cellValue = rowData[col];
                    var cellRef = $"{GetColumnLetter(col)}{row + 1}";
                    object? assignedValue = null;

                    if (cellValue != null)
                    {
                        assignedValue = ConvertJsonElementToValue(cellValue);
                        dynamicSheet.Cells[row, col].Value = assignedValue;
                    }

                    cellAssignments.Add(new CellAssignment
                    {
                        CellReference = cellRef,
                        Row = row,
                        Column = col,
                        OriginalValue = cellValue,
                        OriginalType = cellValue?.GetType().Name ?? "null",
                        AssignedValue = assignedValue,
                        AssignedType = assignedValue?.GetType().Name ?? "null",
                        IsJsonElement = cellValue is JsonElement
                    });
                }
            }
        }

        await _activityPublisher.PublishAsync("dynamic_spreadsheet.created", new
        {
            dataRows,
            dataColumns,
            cellAssignmentCount = cellAssignments.Count,
            timestamp = DateTime.UtcNow
        });

        return new DynamicSpreadsheetResult
        {
            Workbook = dynamicWorkbook,
            Worksheet = dynamicSheet,
            DataRows = dataRows,
            DataColumns = dataColumns,
            CellAssignments = cellAssignments
        };
    }

    /// <summary>
    /// Executes a formula on a workbook and returns the result
    /// </summary>
    public async Task<FormulaExecutionResult> ExecuteFormulaAsync(
        Workbook workbook,
        ExecutionPlanDto executionPlan,
        CancellationToken cancellationToken = default)
    {
        var worksheet = workbook.Worksheets[0];
        int formulaRow = executionPlan.ArtifactsFormatted?.Count ?? 0;
        var formulaCell = worksheet.Cells[formulaRow + 1, 0];
        var formulaCellReference = $"A{formulaRow + 2}";

        var debugInfo = new Dictionary<string, object>
        {
            ["formula"] = executionPlan.Formula,
            ["formulaCellReference"] = formulaCellReference,
            ["formulaRow"] = formulaRow + 1,
            ["formulaColumn"] = 0,
            ["dataRowCount"] = executionPlan.ArtifactsFormatted?.Count ?? 0,
            ["dataColumnCount"] = executionPlan.ArtifactsFormatted?.FirstOrDefault()?.Count ?? 0
        };

        try
        {
            // Set the formula
            formulaCell.Formula = executionPlan.Formula;

            // Calculate the workbook
            workbook.CalculateFormula();

            // Get the result
            var formulaValue = formulaCell.Value;
            
            // Check if the formula operates on date columns
            bool isDateFormula = false;
            if (executionPlan.Formula != null && executionPlan.ArtifactsFormatted != null && executionPlan.ArtifactsFormatted.Count > 0)
            {
                // Check if formula references known date columns
                var formulaUpper = executionPlan.Formula.ToUpper();
                isDateFormula = formulaUpper.Contains("EXDATE") || formulaUpper.Contains("PAYDATE") || 
                               formulaUpper.Contains("DATE");
                
                // Also check if the column header contains date-related keywords
                if (!isDateFormula && executionPlan.ArtifactsFormatted[0].Count > 0)
                {
                    var firstHeader = executionPlan.ArtifactsFormatted[0][0]?.ToString()?.ToUpper() ?? "";
                    isDateFormula = firstHeader.Contains("DATE") || firstHeader == "EXDATE" || firstHeader == "PAYDATE";
                }
            }
            
            var formulaResult = FormatFormulaResult(formulaValue, isDateFormula);

            debugInfo["formulaValue"] = formulaValue ?? "null";
            debugInfo["formulaValueType"] = formulaValue?.GetType().Name ?? "null";
            debugInfo["isError"] = formulaCell.IsErrorValue;
            debugInfo["isDateFormula"] = isDateFormula;

            if (formulaCell.IsErrorValue)
            {
                debugInfo["errorValue"] = formulaCell.Value?.ToString() ?? "Unknown error";
            }

            await _activityPublisher.PublishAsync("formula_execution.completed", debugInfo);

            return new FormulaExecutionResult
            {
                Success = !formulaCell.IsErrorValue,
                Value = formulaValue,
                StringValue = formulaResult,
                Error = formulaCell.IsErrorValue ? formulaCell.Value?.ToString() : null,
                FormulaCellReference = formulaCellReference,
                DebugInfo = debugInfo
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Formula execution failed");

            debugInfo["exception"] = ex.Message;
            debugInfo["stackTrace"] = ex.StackTrace ?? "";

            await _activityPublisher.PublishAsync("formula_execution.failed", debugInfo);

            return new FormulaExecutionResult
            {
                Success = false,
                Value = null,
                StringValue = $"Formula error: {ex.Message}",
                Error = ex.Message,
                FormulaCellReference = formulaCellReference,
                DebugInfo = debugInfo
            };
        }
    }

    /// <summary>
    /// Converts JsonElement to appropriate value type
    /// </summary>
    private object? ConvertJsonElementToValue(object cellValue)
    {
        if (cellValue is JsonElement jsonElement)
        {
            switch (jsonElement.ValueKind)
            {
                case JsonValueKind.Number:
                    return jsonElement.GetDouble();
                case JsonValueKind.String:
                    var stringValue = jsonElement.GetString();
                    
                    // Try to parse as date first
                    if (TryParseDateFromString(stringValue, out var dateValue))
                    {
                        // Convert to Excel date format (OLE Automation date)
                        return dateValue.ToOADate();
                    }
                    
                    // Try to parse as number (handle formatted numbers)
                    if (!string.IsNullOrEmpty(stringValue) && TryParseNumericString(stringValue, out var numValue))
                    {
                        return numValue;
                    }
                    
                    return stringValue;
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return jsonElement.GetBoolean();
                default:
                    return jsonElement.ToString();
            }
        }
        return cellValue;
    }
    
    /// <summary>
    /// Tries to parse various date formats from string
    /// </summary>
    private bool TryParseDateFromString(string? value, out DateTime result)
    {
        result = default;
        if (string.IsNullOrWhiteSpace(value)) return false;
        
        // Common date formats to try
        string[] dateFormats = new[]
        {
            "dd/MM/yyyy", "d/M/yyyy", "dd/MM/yy", "d/M/yy",
            "MM/dd/yyyy", "M/d/yyyy", "MM/dd/yy", "M/d/yy",
            "yyyy-MM-dd", "yyyy/MM/dd",
            "dd-MM-yyyy", "d-M-yyyy",
            "MMM dd, yyyy", "MMMM dd, yyyy",
            "dd MMM yyyy", "dd MMMM yyyy"
        };
        
        foreach (var format in dateFormats)
        {
            if (DateTime.TryParseExact(value, format, 
                System.Globalization.CultureInfo.InvariantCulture, 
                System.Globalization.DateTimeStyles.None, out result))
            {
                return true;
            }
        }
        
        // Try general parsing as last resort
        return DateTime.TryParse(value, out result);
    }
    
    /// <summary>
    /// Formats formula results appropriately based on type
    /// </summary>
    private string FormatFormulaResult(object? value, bool isDateFormula = false)
    {
        if (value == null) return "No result";
        
        switch (value)
        {
            case double d:
                // Handle special cases
                if (double.IsNaN(d)) return "NaN";
                if (double.IsInfinity(d)) return "Infinity";
                
                // If this is a date formula (MIN/MAX on date columns), convert OLE date to ISO format
                if (isDateFormula && d >= 1 && d <= 100000)
                {
                    try
                    {
                        var dateValue = DateTime.FromOADate(d);
                        if (dateValue.Year >= 1900 && dateValue.Year <= 2100)
                        {
                            return dateValue.ToString("yyyy-MM-dd");
                        }
                    }
                    catch
                    {
                        // Not a valid date, continue with numeric formatting
                    }
                }
                
                // Format numbers appropriately
                var absValue = Math.Abs(d);
                
                // For very small numbers close to zero
                if (absValue < 0.0001 && absValue > 0)
                    return d.ToString("E4");
                
                // For percentages or ratios (typically between 0 and 100)
                if (absValue >= 0 && absValue <= 100)
                {
                    // Check if it has meaningful decimal places
                    var rounded = Math.Round(d, 4);
                    if (Math.Abs(rounded - Math.Round(rounded, 2)) < 0.0001)
                        return Math.Round(d, 2).ToString("F2"); // 2 decimal places
                    else
                        return Math.Round(d, 4).ToString("F4"); // 4 decimal places
                }
                
                // For larger numbers
                if (absValue >= 1000000)
                {
                    // For very large variance values (>100M), use exact value with proper decimal
                    if (absValue >= 100000000)
                    {
                        // Check if we need decimal places
                        var rounded = Math.Round(d, 1);
                        if (Math.Abs(rounded - Math.Round(rounded, 0)) < 0.01)
                            return rounded.ToString("F0"); // No decimal places
                        else
                            return rounded.ToString("F1"); // 1 decimal place like "291839462.7"
                    }
                    else
                    {
                        // Regular large numbers with thousands separator
                        var rounded = Math.Round(d, 2);
                        if (rounded % 1 == 0)
                            return rounded.ToString("N0"); // No decimal places, with thousands separator
                        else
                            return rounded.ToString("N2").Replace(",", " "); // 2 decimal places with space as thousands separator
                    }
                }
                else if (absValue >= 1000)
                {
                    // For thousands, use space as separator like "31 683.10"
                    var rounded = Math.Round(d, 2);
                    var formatted = rounded.ToString("N2");
                    return formatted.Replace(",", " ");
                }
                else if (absValue >= 100)
                {
                    // Ensure proper rounding to 2 decimal places
                    return Math.Round(d, 2).ToString("F2");
                }
                else
                {
                    // Small numbers - use 4 decimal places but round first
                    return Math.Round(d, 4).ToString("F4");
                }
                    
            case float f:
                return FormatFormulaResult((double)f, isDateFormula);
                
            case decimal dec:
                return FormatFormulaResult((double)dec, isDateFormula);
                
            case int i:
                return i.ToString("N0");
                
            case long l:
                return l.ToString("N0");
                
            case DateTime dt:
                return dt.ToString("yyyy-MM-dd");
                
            case bool b:
                return b ? "TRUE" : "FALSE";
                
            default:
                return value.ToString() ?? "No result";
        }
    }

    #region Helper Methods

    private bool EvaluateFilter(object cellValue, FilterCriteria filter)
    {
        switch (filter.Operator.ToLower())
        {
            case "equals":
                return CompareEquals(cellValue, filter.Value);

            case "contains":
                return cellValue?.ToString()?.Contains(filter.Value, StringComparison.OrdinalIgnoreCase) ?? false;

            case ">":
            case "<":
            case ">=":
            case "<=":
                if (TryParseNumeric(cellValue, out var numVal) &&
                    double.TryParse(filter.Value, out var filterNum))
                {
                    return filter.Operator switch
                    {
                        ">" => numVal > filterNum,
                        "<" => numVal < filterNum,
                        ">=" => numVal >= filterNum,
                        "<=" => numVal <= filterNum,
                        _ => false
                    };
                }

                break;

            case "date>":
            case "date<":
            case "date>=":
            case "date<=":
                var dateVal = TryParseDate(cellValue);
                if (dateVal.HasValue && DateTime.TryParse(filter.Value, out var filterDate))
                {
                    return filter.Operator switch
                    {
                        "date>" => dateVal.Value > filterDate,
                        "date<" => dateVal.Value < filterDate,
                        "date>=" => dateVal.Value >= filterDate,
                        "date<=" => dateVal.Value <= filterDate,
                        _ => false
                    };
                }

                break;
        }

        return false;
    }

    private bool CompareEquals(object cellValue, string filterValue)
    {
        if (cellValue == null) return string.IsNullOrEmpty(filterValue);

        var cellStr = cellValue.ToString() ?? "";

        // Try numeric comparison first
        if (TryParseNumeric(cellValue, out var cellNum) &&
            double.TryParse(filterValue, out var filterNum))
        {
            return Math.Abs(cellNum - filterNum) < 0.0001;
        }

        // Try date comparison
        var cellDate = TryParseDate(cellValue);
        if (cellDate.HasValue && DateTime.TryParse(filterValue, out var filterDate))
        {
            return cellDate.Value.Date == filterDate.Date;
        }

        // Fall back to string comparison
        return cellStr.Equals(filterValue, StringComparison.OrdinalIgnoreCase);
    }

    private bool TryParseNumeric(object cellValue, out double result)
    {
        result = 0;

        if (cellValue == null) return false;

        return cellValue switch
        {
            double d => (result = d, true).Item2,
            int i => (result = i, true).Item2,
            decimal dec => (result = (double)dec, true).Item2,
            string s => TryParseNumericString(s, out result),
            _ => double.TryParse(cellValue.ToString(), out result)
        };
    }

    private bool TryParseNumericString(string s, out double result)
    {
        result = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;

        // Clean the string
        s = s.Trim();

        // Check for accounting format (negative numbers in parentheses)
        bool isNegative = false;
        if (s.StartsWith("(") && s.EndsWith(")"))
        {
            isNegative = true;
            s = s.Substring(1, s.Length - 2).Trim();
        }
        
        // Check for trailing minus (accounting style)
        if (s.EndsWith("-"))
        {
            isNegative = true;
            s = s.Substring(0, s.Length - 1).Trim();
        }

        // Remove currency symbols and thousands separators
        s = s.Replace("$", "").Replace(",", "").Replace("€", "").Replace("£", "").Trim();
        
        // Handle leading minus sign (preserve it)
        if (s.StartsWith("-"))
        {
            isNegative = true;
            s = s.Substring(1).Trim();
        }

        // Try to parse the cleaned string
        if (double.TryParse(s, out result))
        {
            if (isNegative)
            {
                result = -result;
            }

            return true;
        }

        return false;
    }

    private DateTime? TryParseDate(object cellValue)
    {
        if (cellValue == null!) return null;

        if (cellValue is DateTime dt) return dt;

        if (cellValue is double oaDate)
        {
            try
            {
                return DateTime.FromOADate(oaDate);
            }
            catch
            {
                return null;
            }
        }

        if (DateTime.TryParse(cellValue.ToString(), out var parsed))
            return parsed;

        return null;
    }

    #endregion

    #region DTOs

    /// <summary>
    /// Sample data extracted from worksheet
    /// </summary>
    private class SampleData
    {
        public Dictionary<string, List<object>> Data { get; set; } = new();
    }

    /// <summary>
    /// Result from LLM analysis of a sample
    /// </summary>
    private class LlmAnalysisResult
    {
        public List<string> NewHeadersFound { get; init; } = [];
        public bool HasSufficientContext { get; init; }
        public string ArtifactContent { get; init; } = "";
        public List<string> NeededData { get; init; } = [];
        public string UserIntentWithContext { get; init; } = "";
        public List<string> RelevantPatterns { get; init; } = [];
        public string SuggestedFormula { get; init; } = "";
        public double ConfidenceScore { get; init; }
    }

    /// <summary>
    /// Format detection result from LLM
    /// </summary>
    private class FormatDetectionResult
    {
        public string Format { get; init; } = "Unknown";
        public double Confidence { get; init; }
        public string Reasoning { get; init; } = "";
        public string HeaderLocation { get; init; } = "";
    }

    /// <summary>
    /// Header extraction result from LLM
    /// </summary>
    private class HeaderExtractionResult
    {
        public int HeaderRowIndex { get; init; }
        public List<string> Headers { get; init; } = [];
        public bool MultiRowHeaders { get; init; }
        public double Confidence { get; init; }
    }

    /// <summary>
    /// Final analysis response from LLM
    /// </summary>
    private class FinalAnalysisResponse
    {
        public List<string> ColumnsNeeded { get; init; } = [];
        public List<FilterResponse> Filters { get; init; } = [];
        public string AggregationType { get; init; } = "";
        public string? GroupBy { get; init; }
        public bool RequiresCalculation { get; init; }
        public List<string> CalculationSteps { get; init; } = [];
        public bool RequiresFullDataset { get; init; }
        public string UserIntentWithContext { get; init; } = "";
        public string Artifact { get; init; } = "";
        public List<object> ContextSnapshots { get; init; } = [];
    }

    private class FilterResponse
    {
        public string Column { get; set; } = "";
        public string Operator { get; set; } = "";
        public string Value { get; set; } = "";
    }

    /// <summary>
    /// Result from batch analysis
    /// </summary>
    private class BatchAnalysisResult
    {
        public string NewArtifacts { get; set; } = "";
        public bool ContinueSnapshot { get; set; }
        public string UserIntentWithContext { get; set; } = "";
        public string Reasoning { get; set; } = "";
    }

    /// <summary>
    /// Execution plan response from LLM
    /// </summary>
    private class ExecutionPlanResponse
    {
        public bool NeedRunFormula { get; set; }
        public List<List<object>> ArtifactsFormatted { get; set; } = new();
        public string Formula { get; set; } = "";
        public string SimpleAnswer { get; set; } = "";
        public string Reasoning { get; set; } = "";
    }

    #endregion
}