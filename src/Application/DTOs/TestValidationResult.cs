namespace SpreadsheetCLI.Application.DTOs;

/// <summary>
/// Result of test validation using LLM
/// </summary>
public class TestValidationResult
{
    /// <summary>
    /// Whether the test passed based on LLM evaluation
    /// </summary>
    public bool IsCorrect { get; set; }
    
    /// <summary>
    /// The actual answer extracted from the output
    /// </summary>
    public string ExtractedAnswer { get; set; } = "";
    
    /// <summary>
    /// Explanation of why the test passed or failed
    /// </summary>
    public string Explanation { get; set; } = "";
    
    /// <summary>
    /// Confidence level of the validation (0-1)
    /// </summary>
    public double Confidence { get; set; }
    
    /// <summary>
    /// Where the answer was found (e.g., "Answer field", "Reasoning field", "MachineAnswer field")
    /// </summary>
    public string AnswerLocation { get; set; } = "";
}

/// <summary>
/// Request for test validation
/// </summary>
public class TestValidationRequest
{
    /// <summary>
    /// The question that was asked
    /// </summary>
    public string Question { get; set; } = "";
    
    /// <summary>
    /// The expected answer from ground truth
    /// </summary>
    public string ExpectedAnswer { get; set; } = "";
    
    /// <summary>
    /// The full JSON output from the CLI
    /// </summary>
    public string ActualOutput { get; set; } = "";
    
    /// <summary>
    /// Optional reasoning field if already extracted
    /// </summary>
    public string? Reasoning { get; set; }
    
    /// <summary>
    /// Optional answer field if already extracted
    /// </summary>
    public string? Answer { get; set; }
    
    /// <summary>
    /// Optional machine answer field if already extracted
    /// </summary>
    public string? MachineAnswer { get; set; }
}