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

namespace SpreadsheetCLI.Mock;

/// <summary>
/// Unified service for spreadsheet analysis with AI-powered query processing
/// </summary>
public class SpreadsheetService : ISpreadsheetService
{
    private readonly ILogger<SpreadsheetService> _logger;
    private readonly IChatCompletionService _chatCompletion;
    
    private const int SampleSize = 50;
    private const int MaxIterations = 10;
    private const int MaxInitSampleColumns = 50;

    public SpreadsheetService(
        ILogger<SpreadsheetService> logger,
        IChatCompletionService chatCompletion)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _chatCompletion = chatCompletion ?? throw new ArgumentNullException(nameof(chatCompletion));
    }

    /// <summary>
    /// Analyzes a natural language query against spreadsheet data
    /// </summary>
    public async Task<QueryAnalysisResult> AnalyzeQueryAsync(
        string query,
        DocumentMetadata metadata,
        Worksheet worksheet,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Starting dynamic query analysis: {Query}", query);
        
        var artifacts = new List<string>();
        var currentRowIndex = metadata.DataStartRow;
        var iterationCount = 0;
        var continueAnalysis = true;
        var userIntentWithContext = "";

        // Iterative analysis loop
        while (continueAnalysis && iterationCount < MaxIterations && currentRowIndex <= worksheet.Cells.MaxRow)
        {
            iterationCount++;

            // Build markdown table with current batch
            var markdownTable = BuildMarkdownTable(worksheet, metadata, currentRowIndex, SampleSize);

            // Analyze batch with LLM
            var batchResult = await AnalyzeBatchWithLlmAsync(
                query, metadata, markdownTable, artifacts, iterationCount, cancellationToken);

            if (!string.IsNullOrWhiteSpace(batchResult.NewArtifacts))
                artifacts.Add(batchResult.NewArtifacts);

            if (!string.IsNullOrWhiteSpace(batchResult.UserIntentWithContext))
                userIntentWithContext = batchResult.UserIntentWithContext;

            continueAnalysis = batchResult.ContinueSnapshot;
            currentRowIndex += SampleSize;
        }

        // Generate execution plan from artifacts
        var combinedArtifacts = string.Join("\n\n---\n\n", artifacts);
        return await GenerateExecutionPlanFromArtifactsAsync(
            query, userIntentWithContext, combinedArtifacts, metadata, cancellationToken);
    }

    /// <summary>
    /// Detects document format using stratified sampling
    /// </summary>
    public async Task<DocumentFormat> DetectDocumentFormatAsync(
        Worksheet worksheet,
        CancellationToken cancellationToken = default)
    {
        var samples = GatherStratifiedSamples(worksheet);
        
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
            """;

        var responseFormat = CreateFormatDetectionResponseFormat();
        var response = await GetLlmResponseAsync(prompt, responseFormat, "gpt-4o-mini", cancellationToken);
        
        var result = JsonSerializer.Deserialize<FormatDetectionResult>(response);
        return Enum.Parse<DocumentFormat>(result?.Format ?? "Unknown");
    }

    /// <summary>
    /// Extracts document metadata
    /// </summary>
    public async Task<DocumentMetadata> ExtractDocumentMetadataAsync(
        Worksheet worksheet,
        DocumentFormat format,
        List<HeaderInfo> headers,
        CancellationToken cancellationToken = default)
    {
        var headerRowIndex = headers.Any() ? headers.First().RowIndex : 0;
        var dataStartRow = headerRowIndex + 1;
        var totalRows = worksheet.Cells.MaxRow + 1;
        var dataRowCount = Math.Max(0, totalRows - dataStartRow);

        var metadata = new DocumentMetadata
        {
            Format = format,
            TotalRows = totalRows,
            TotalColumns = worksheet.Cells.MaxColumn + 1,
            DataStartRow = dataStartRow,
            DataRowCount = dataRowCount,
            Headers = headers.Select(h => h.Name).ToList()
        };

        // Detect data types
        metadata.DataTypes = await DetectDataTypesAsync(worksheet, metadata, headers, cancellationToken);
        
        return metadata;
    }

    /// <summary>
    /// Extracts headers from worksheet
    /// </summary>
    public List<HeaderInfo> ExtractHeaders(Worksheet worksheet)
    {
        return Task.Run(async () => 
            await ExtractHeadersWithLlmAsync(worksheet, CancellationToken.None)
        ).GetAwaiter().GetResult();
    }

    /// <summary>
    /// Creates dynamic spreadsheet from execution plan
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
                    object? assignedValue = ConvertJsonElementToValue(cellValue);
                    
                    if (assignedValue != null)
                        dynamicSheet.Cells[row, col].Value = assignedValue;

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
    /// Executes formula on workbook
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

        try
        {
            formulaCell.Formula = executionPlan.Formula;
            workbook.CalculateFormula();
            
            var formulaValue = formulaCell.Value;
            bool isDateFormula = CheckIfDateFormula(executionPlan);
            var formulaResult = FormatFormulaResult(formulaValue, isDateFormula);

            return new FormulaExecutionResult
            {
                Success = !formulaCell.IsErrorValue,
                Value = formulaValue,
                StringValue = formulaResult,
                Error = formulaCell.IsErrorValue ? formulaCell.Value?.ToString() : null,
                FormulaCellReference = formulaCellReference,
                DebugInfo = new Dictionary<string, object>
                {
                    ["formula"] = executionPlan.Formula,
                    ["formulaCellReference"] = formulaCellReference,
                    ["formulaValue"] = formulaValue ?? "null",
                    ["formulaValueType"] = formulaValue?.GetType().Name ?? "null"
                }
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Formula execution failed");
            return new FormulaExecutionResult
            {
                Success = false,
                StringValue = $"Formula error: {ex.Message}",
                Error = ex.Message,
                FormulaCellReference = formulaCellReference
            };
        }
    }

    #region Private Helper Methods

    private List<string> GatherStratifiedSamples(Worksheet worksheet)
    {
        var maxRow = worksheet.Cells.MaxRow;
        var maxCol = worksheet.Cells.MaxColumn;
        var samples = new List<string>();

        // First 50 rows
        for (int row = 0; row < Math.Min(SampleSize, maxRow + 1); row++)
        {
            samples.Add(GetRowSample(worksheet, row, Math.Min(10, maxCol)));
        }

        // Middle 50 rows
        if (maxRow > SampleSize * 2)
        {
            var middleStart = maxRow / 2 - SampleSize / 2;
            for (int row = middleStart; row < middleStart + SampleSize; row++)
            {
                samples.Add(GetRowSample(worksheet, row, Math.Min(10, maxCol)));
            }
        }

        // Last 50 rows
        if (maxRow > SampleSize)
        {
            for (int row = Math.Max(0, maxRow - SampleSize + 1); row <= maxRow; row++)
            {
                samples.Add(GetRowSample(worksheet, row, Math.Min(10, maxCol)));
            }
        }

        return samples;
    }

    private string GetRowSample(Worksheet worksheet, int row, int maxCol)
    {
        var rowData = new List<string>();
        for (int col = 0; col <= maxCol; col++)
        {
            rowData.Add(worksheet.Cells[row, col].StringValue ?? "");
        }
        return string.Join("|", rowData);
    }

    private async Task<List<HeaderInfo>> ExtractHeadersWithLlmAsync(
        Worksheet worksheet,
        CancellationToken cancellationToken)
    {
        var maxRows = Math.Min(MaxInitSampleColumns, worksheet.Cells.MaxRow + 1);
        var markdownTable = CreateMarkdownTableWithRealIndices(worksheet, 0, maxRows - 1);

        var prompt = $"""
            Analyze this spreadsheet data and identify the column headers.

            IMPORTANT: The row numbers shown are the ACTUAL row indices from the Excel file.
            Total rows in document: {worksheet.Cells.MaxRow + 1}

            {markdownTable}

            Identify:
            1. Which row contains the headers
            2. Extract all column headers from that row
            3. Return the ACTUAL row index where headers are found
            """;

        var responseFormat = CreateHeaderExtractionResponseFormat();
        var response = await GetLlmResponseAsync(prompt, responseFormat, "gpt-4o-mini", cancellationToken);
        
        var result = JsonSerializer.Deserialize<HeaderExtractionResult>(response);
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

    private string CreateMarkdownTableWithRealIndices(Worksheet worksheet, int startRow, int endRow)
    {
        var sb = new StringBuilder();
        var actualMaxRow = worksheet.Cells.MaxRow;
        var actualMaxCol = worksheet.Cells.MaxColumn;

        startRow = Math.Max(0, Math.Min(startRow, actualMaxRow));
        endRow = Math.Max(startRow, Math.Min(endRow, actualMaxRow));

        if (actualMaxRow < 0 || actualMaxCol < 0)
            return "*Empty worksheet*";

        var colsToShow = Math.Min(10, actualMaxCol);

        // Table header
        sb.Append("| Row # |");
        for (int col = 0; col <= colsToShow; col++)
            sb.Append($" Col {col} |");
        sb.AppendLine();

        // Separator
        sb.Append("|-------|");
        for (int col = 0; col <= colsToShow; col++)
            sb.Append("--------|");
        sb.AppendLine();

        // Data rows
        for (int row = startRow; row <= endRow; row++)
        {
            sb.Append($"| {row} |");
            for (int col = 0; col <= colsToShow; col++)
            {
                var cellValue = worksheet.Cells[row, col].StringValue ?? "";
                cellValue = cellValue.Replace("|", "\\|").Replace("\n", " ").Trim();
                if (cellValue.Length > 30)
                    cellValue = cellValue.Substring(0, 27) + "...";
                sb.Append($" {cellValue} |");
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    private string BuildMarkdownTable(Worksheet worksheet, DocumentMetadata metadata, int startRow, int rowCount)
    {
        var sb = new StringBuilder();
        var actualMaxRow = worksheet.Cells.MaxRow;
        var actualMaxCol = worksheet.Cells.MaxColumn;

        if (startRow > actualMaxRow)
        {
            sb.AppendLine("### No more data to analyze");
            return sb.ToString();
        }

        var endRow = Math.Min(startRow + rowCount - 1, actualMaxRow);
        var colsToShow = Math.Min(actualMaxCol, 20);

        sb.AppendLine("### Data Sample");
        sb.AppendLine($"*Showing rows {startRow} to {endRow}*");
        sb.AppendLine();

        // Header row
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
        sb.AppendLine();

        // Separator
        sb.Append("|:---:|");
        for (int col = 0; col <= colsToShow; col++)
            sb.Append(":------:|");
        sb.AppendLine();

        // Data rows
        for (int row = startRow; row <= endRow && row <= actualMaxRow; row++)
        {
            sb.Append($"| **{row}** | ");
            for (int col = 0; col <= colsToShow; col++)
            {
                var cellValue = worksheet.Cells[row, col].StringValue ?? "";
                cellValue = cellValue.Replace("|", "\\|").Replace("\n", " ").Trim();
                if (cellValue.Length > 50)
                    cellValue = cellValue.Substring(0, 47) + "...";
                if (string.IsNullOrWhiteSpace(cellValue))
                    cellValue = "_empty_";
                sb.Append($"{cellValue} | ");
            }
            sb.AppendLine();
        }

        // Dataset info
        sb.AppendLine();
        sb.AppendLine("### 📊 Dataset Info");
        sb.AppendLine($"- Headers at row: **{metadata.DataStartRow - 1}**");
        sb.AppendLine($"- Data range: rows **{metadata.DataStartRow}** to **{metadata.TotalRows - 1}**");
        sb.AppendLine($"- Total data rows: **{metadata.DataRowCount:N0}**");

        return sb.ToString();
    }

    private async Task<BatchAnalysisResult> AnalyzeBatchWithLlmAsync(
        string query,
        DocumentMetadata metadata,
        string markdownTable,
        List<string> previousArtifacts,
        int iteration,
        CancellationToken cancellationToken)
    {
        var prompt = $"""
            You are analyzing spreadsheet data to answer this query: {query}

            Document metadata:
            - Total rows in file: {metadata.TotalRows}
            - Headers at row: {metadata.DataStartRow - 1}
            - Data rows: {metadata.DataRowCount}
            - Headers: {string.Join(", ", metadata.Headers)}

            This is iteration {iteration} of the analysis.

            Current data sample:
            {markdownTable}

            Instructions:
            1. Extract any data relevant to the query
            2. Store extracted data as "artifacts"
            3. Determine if you need to continue analyzing more rows
            4. Provide the user intent with context
            """;

        var responseFormat = CreateBatchAnalysisResponseFormat();
        var response = await GetLlmResponseAsync(prompt, responseFormat, "gpt-4o-mini", cancellationToken);
        
        return JsonSerializer.Deserialize<BatchAnalysisResult>(response) ?? new BatchAnalysisResult();
    }

    private async Task<QueryAnalysisResult> GenerateExecutionPlanFromArtifactsAsync(
        string query,
        string userIntentWithContext,
        string combinedArtifacts,
        DocumentMetadata metadata,
        CancellationToken cancellationToken)
    {
        var prompt = $"""
            Generate an execution plan for this query: {query}

            User intent: {userIntentWithContext}

            Use Excel standard functions with correct syntax.

            Instructions:
            1. Create ArtifactsFormatted as a 2D array
            2. Generate appropriate Excel formula
            3. Provide reasoning

            Collected artifacts:
            {combinedArtifacts}

            Headers: {string.Join(", ", metadata.Headers)}
            """;

        var responseFormat = CreateExecutionPlanResponseFormat();
        var response = await GetLlmResponseAsync(prompt, responseFormat, "gpt-4o-mini-high", cancellationToken);
        
        var result = JsonSerializer.Deserialize<ExecutionPlanResponse>(response);

        return new QueryAnalysisResult
        {
            ColumnsNeeded = ExtractColumnsFromArtifacts(result?.ArtifactsFormatted),
            RequiresCalculation = result?.NeedRunFormula ?? false,
            CalculationSteps = new List<string> { result?.Formula ?? "" },
            UserIntentWithContext = userIntentWithContext,
            Artifact = JsonSerializer.Serialize(new
            {
                ExecutionPlan = result,
                OriginalArtifacts = combinedArtifacts
            })
        };
    }

    private List<string> ExtractColumnsFromArtifacts(List<List<object>>? artifacts)
    {
        if (artifacts == null || artifacts.Count == 0)
            return new List<string>();
        return artifacts[0]?.Select(h => h?.ToString() ?? "").ToList() ?? new List<string>();
    }

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

    private object? ConvertJsonElementToValue(object? cellValue)
    {
        if (cellValue is JsonElement jsonElement)
        {
            switch (jsonElement.ValueKind)
            {
                case JsonValueKind.Number:
                    return jsonElement.GetDouble();
                case JsonValueKind.String:
                    var stringValue = jsonElement.GetString();
                    if (TryParseDateFromString(stringValue, out var dateValue))
                        return dateValue.ToOADate();
                    if (!string.IsNullOrEmpty(stringValue) && TryParseNumericString(stringValue, out var numValue))
                        return numValue;
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

    private bool TryParseDateFromString(string? value, out DateTime result)
    {
        result = default;
        if (string.IsNullOrWhiteSpace(value)) return false;
        
        string[] dateFormats = new[]
        {
            "dd/MM/yyyy", "d/M/yyyy", "MM/dd/yyyy", "M/d/yyyy",
            "yyyy-MM-dd", "yyyy/MM/dd", "dd-MM-yyyy", "dd MMM yyyy"
        };
        
        foreach (var format in dateFormats)
        {
            if (DateTime.TryParseExact(value, format, 
                System.Globalization.CultureInfo.InvariantCulture, 
                System.Globalization.DateTimeStyles.None, out result))
                return true;
        }
        
        return DateTime.TryParse(value, out result);
    }

    private bool TryParseNumericString(string s, out double result)
    {
        result = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;

        s = s.Trim();
        bool isNegative = false;
        
        if (s.StartsWith("(") && s.EndsWith(")"))
        {
            isNegative = true;
            s = s.Substring(1, s.Length - 2).Trim();
        }
        
        s = s.Replace("$", "").Replace(",", "").Replace("€", "").Replace("£", "").Trim();
        
        if (double.TryParse(s, out result))
        {
            if (isNegative) result = -result;
            return true;
        }

        return false;
    }

    private bool CheckIfDateFormula(ExecutionPlanDto executionPlan)
    {
        if (executionPlan.Formula == null || executionPlan.ArtifactsFormatted == null) 
            return false;
            
        var formulaUpper = executionPlan.Formula.ToUpper();
        return formulaUpper.Contains("DATE") || formulaUpper.Contains("EXDATE") || formulaUpper.Contains("PAYDATE");
    }

    private string FormatFormulaResult(object? value, bool isDateFormula = false)
    {
        if (value == null) return "No result";
        
        switch (value)
        {
            case double d:
                if (double.IsNaN(d)) return "NaN";
                if (double.IsInfinity(d)) return "Infinity";
                
                if (isDateFormula && d >= 1 && d <= 100000)
                {
                    try
                    {
                        var dateValue = DateTime.FromOADate(d);
                        if (dateValue.Year >= 1900 && dateValue.Year <= 2100)
                            return dateValue.ToString("yyyy-MM-dd");
                    }
                    catch { }
                }
                
                return FormatNumericValue(d);
                
            case float f:
                return FormatFormulaResult((double)f, isDateFormula);
                
            case decimal dec:
                return FormatFormulaResult((double)dec, isDateFormula);
                
            case int i:
                return i.ToString("N0");
                
            case DateTime dt:
                return dt.ToString("yyyy-MM-dd");
                
            case bool b:
                return b ? "TRUE" : "FALSE";
                
            default:
                return value.ToString() ?? "No result";
        }
    }

    private string FormatNumericValue(double d)
    {
        var absValue = Math.Abs(d);
        
        if (absValue < 0.0001 && absValue > 0)
            return d.ToString("E4");
        
        if (absValue >= 0 && absValue <= 100)
            return Math.Round(d, 2).ToString("F2");
        
        if (absValue >= 100000000)
            return Math.Round(d, 1).ToString("F1");
        
        if (absValue >= 1000)
            return Math.Round(d, 2).ToString("F2");
        
        return Math.Round(d, 4).ToString("F4");
    }

    private async Task<Dictionary<string, string>> DetectDataTypesAsync(
        Worksheet worksheet,
        DocumentMetadata metadata,
        List<HeaderInfo> headers,
        CancellationToken cancellationToken)
    {
        var dataTypes = new Dictionary<string, string>();
        
        foreach (var header in headers)
        {
            var colIndex = headers.IndexOf(header);
            var sampleValues = new List<object>();
            
            for (int row = metadata.DataStartRow; row < Math.Min(metadata.DataStartRow + 100, metadata.TotalRows); row++)
            {
                var value = worksheet.Cells[row, colIndex].Value;
                if (value != null) sampleValues.Add(value);
            }
            
            dataTypes[header.Name] = DetermineDataType(sampleValues);
        }
        
        return dataTypes;
    }

    private string DetermineDataType(List<object> values)
    {
        if (!values.Any()) return "unknown";
        
        int numericCount = 0, dateCount = 0, textCount = 0;
        
        foreach (var value in values)
        {
            if (value is double || value is int || value is decimal)
                numericCount++;
            else if (value is DateTime)
                dateCount++;
            else
                textCount++;
        }
        
        var total = values.Count;
        if (numericCount > total * 0.8) return "numeric";
        if (dateCount > total * 0.8) return "date";
        return "text";
    }

    #endregion

    #region LLM Communication

    private async Task<string> GetLlmResponseAsync(
        string prompt,
        ChatResponseFormat responseFormat,
        string modelId,
        CancellationToken cancellationToken)
    {
        var settings = new OpenAIPromptExecutionSettings
        {
            ModelId = modelId,
            ResponseFormat = responseFormat,
            Temperature = 0.1
        };

        var chatHistory = new ChatHistory();
        chatHistory.AddMessage(AuthorRole.User, prompt);

        var response = await _chatCompletion.GetChatMessageContentsAsync(
            chatHistory, settings, cancellationToken: cancellationToken);

        return response[0].Content ?? "{}";
    }

    private ChatResponseFormat CreateFormatDetectionResponseFormat()
    {
        return ChatResponseFormat.CreateJsonSchemaFormat(
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
                        "Reasoning": { "type": "string" }
                    },
                    "required": ["Format", "Confidence", "Reasoning"],
                    "additionalProperties": false
                }
                """),
            jsonSchemaIsStrict: true
        );
    }

    private ChatResponseFormat CreateHeaderExtractionResponseFormat()
    {
        return ChatResponseFormat.CreateJsonSchemaFormat(
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
                        "Confidence": { "type": "number" }
                    },
                    "required": ["HeaderRowIndex", "Headers", "Confidence"],
                    "additionalProperties": false
                }
                """),
            jsonSchemaIsStrict: true
        );
    }

    private ChatResponseFormat CreateBatchAnalysisResponseFormat()
    {
        return ChatResponseFormat.CreateJsonSchemaFormat(
            jsonSchemaFormatName: "batch_analysis",
            jsonSchema: BinaryData.FromString("""
                {
                    "type": "object",
                    "properties": {
                        "NewArtifacts": { "type": "string" },
                        "ContinueSnapshot": { "type": "boolean" },
                        "UserIntentWithContext": { "type": "string" },
                        "Reasoning": { "type": "string" }
                    },
                    "required": ["NewArtifacts", "ContinueSnapshot", "UserIntentWithContext", "Reasoning"],
                    "additionalProperties": false
                }
                """),
            jsonSchemaIsStrict: true
        );
    }

    private ChatResponseFormat CreateExecutionPlanResponseFormat()
    {
        return ChatResponseFormat.CreateJsonSchemaFormat(
            jsonSchemaFormatName: "execution_plan",
            jsonSchema: BinaryData.FromString("""
                {
                    "type": "object",
                    "properties": {
                        "NeedRunFormula": { "type": "boolean" },
                        "ArtifactsFormatted": {
                            "type": "array",
                            "items": {
                                "type": "array",
                                "items": {
                                    "type": ["string", "number", "boolean", "null"]
                                }
                            }
                        },
                        "Formula": { "type": "string" },
                        "SimpleAnswer": { "type": "string" },
                        "Reasoning": { "type": "string" },
                        "MachineAnswer": { "type": "string" },
                        "HumanExplanation": { "type": "string" }
                    },
                    "required": ["NeedRunFormula", "ArtifactsFormatted", "Formula", "SimpleAnswer", "Reasoning", "MachineAnswer", "HumanExplanation"],
                    "additionalProperties": false
                }
                """),
            jsonSchemaIsStrict: true
        );
    }

    #endregion

    #region Response DTOs

    private class FormatDetectionResult
    {
        public string Format { get; set; } = "Unknown";
        public double Confidence { get; set; }
        public string Reasoning { get; set; } = "";
    }

    private class HeaderExtractionResult
    {
        public int HeaderRowIndex { get; set; }
        public List<string> Headers { get; set; } = new();
        public double Confidence { get; set; }
    }

    private class BatchAnalysisResult
    {
        public string NewArtifacts { get; set; } = "";
        public bool ContinueSnapshot { get; set; }
        public string UserIntentWithContext { get; set; } = "";
        public string Reasoning { get; set; } = "";
    }

    private class ExecutionPlanResponse
    {
        public bool NeedRunFormula { get; set; }
        public List<List<object>> ArtifactsFormatted { get; set; } = new();
        public string Formula { get; set; } = "";
        public string SimpleAnswer { get; set; } = "";
        public string Reasoning { get; set; } = "";
        public string MachineAnswer { get; set; } = "";
        public string HumanExplanation { get; set; } = "";
    }

    #endregion
}