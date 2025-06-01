using System.Collections.Generic;

namespace SpreadsheetCLI.Application.DTOs;

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
    
    /// <summary>
    /// Machine-readable numeric answer (no formatting, just the value)
    /// Used for test validation and programmatic access
    /// </summary>
    public string? MachineAnswer { get; set; }
    
    /// <summary>
    /// Human-readable narrative explanation
    /// Can include context like "Group X has the highest variance"
    /// </summary>
    public string? HumanExplanation { get; set; }
}