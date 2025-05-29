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
}