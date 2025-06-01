using System;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using OpenAI.Chat;
using SpreadsheetCLI.Application.DTOs;
using SpreadsheetCLI.Application.Interfaces;

namespace SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Services;

/// <summary>
/// Service for validating test results using LLM with structured outputs
/// </summary>
public class TestResultValidationService : ITestResultValidationService
{
    private readonly ILogger<TestResultValidationService> _logger;
    private readonly IChatCompletionService _chatCompletion;

    public TestResultValidationService(
        ILogger<TestResultValidationService> logger,
        IChatCompletionService chatCompletion)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _chatCompletion = chatCompletion ?? throw new ArgumentNullException(nameof(chatCompletion));
    }

    public async Task<TestValidationResult> ValidateTestResultAsync(
        TestValidationRequest request,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Validating test result for question: {Question}", request.Question);

        var prompt = $"""
            You are a test validator for a spreadsheet analysis CLI. Your job is to determine if the actual output correctly answers the given question.

            Question: {request.Question}
            Expected Answer: {request.ExpectedAnswer}

            The CLI output may contain the answer in various fields:
            - Answer: Direct answer field
            - MachineAnswer: Machine-readable answer (numeric values, entity names)
            - Reasoning: Explanation that may contain the answer
            - SimpleAnswer: Alternative answer field
            - HumanExplanation: Human-readable explanation

            Actual CLI Output:
            {request.ActualOutput}

            VALIDATION RULES:
            1. Numeric Comparison:
               - Allow for small rounding differences (±0.01 for decimals)
               - Ignore formatting differences (commas, dollar signs, percent signs)
               - "12977.52" matches "12,977.52" or "$12977.52"
               - For percentages: "48.69" matches "48.69%" 
               
            2. Text Comparison:
               - Case-insensitive matching
               - "MSFT" matches "msft" or "Msft"
               
            3. Answer Extraction Priority:
               - First check MachineAnswer field (most reliable for numeric/entity answers)
               - Then check Answer and SimpleAnswer fields
               - Finally search in Reasoning field for answer patterns
               - Look for patterns like "The answer is X", "total is X", "average is X"
               
            4. Special Cases:
               - For "Which X has highest/lowest Y" questions: Accept either just the X value or the Y value
               - For percentage questions: Accept with or without % sign
               - For count questions: Look for "X rows", "X records", "X unique values"
               
            IMPORTANT: Be flexible in extraction but strict in validation. The answer must be semantically correct.

            Extract the actual answer from the output and determine if it matches the expected answer.
            """;

        var responseFormat = ChatResponseFormat.CreateJsonSchemaFormat(
            jsonSchemaFormatName: "test_validation",
            jsonSchema: BinaryData.FromString("""
                {
                    "type": "object",
                    "properties": {
                        "IsCorrect": {
                            "type": "boolean",
                            "description": "Whether the actual output correctly answers the question"
                        },
                        "ExtractedAnswer": {
                            "type": "string",
                            "description": "The answer extracted from the actual output"
                        },
                        "Explanation": {
                            "type": "string",
                            "description": "Brief explanation of why the test passed or failed"
                        },
                        "Confidence": {
                            "type": "number",
                            "description": "Confidence level of the validation (0-1)",
                            "minimum": 0,
                            "maximum": 1
                        },
                        "AnswerLocation": {
                            "type": "string",
                            "description": "Where the answer was found (e.g., 'MachineAnswer field', 'Reasoning field')"
                        }
                    },
                    "required": ["IsCorrect", "ExtractedAnswer", "Explanation", "Confidence", "AnswerLocation"],
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

        try
        {
            var response = await _chatCompletion.GetChatMessageContentsAsync(
                chatHistory, settings, cancellationToken: cancellationToken);

            var result = JsonSerializer.Deserialize<TestValidationResult>(response[0].Content ?? "{}");
            
            if (result == null)
            {
                throw new InvalidOperationException("Failed to parse validation response");
            }

            _logger.LogInformation(
                "Validation completed - IsCorrect: {IsCorrect}, ExtractedAnswer: {ExtractedAnswer}, Location: {Location}",
                result.IsCorrect, result.ExtractedAnswer, result.AnswerLocation);

            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during test validation");
            
            // Return a failed validation with error details
            return new TestValidationResult
            {
                IsCorrect = false,
                ExtractedAnswer = "Error during validation",
                Explanation = $"Validation failed due to error: {ex.Message}",
                Confidence = 0,
                AnswerLocation = "N/A"
            };
        }
    }
}