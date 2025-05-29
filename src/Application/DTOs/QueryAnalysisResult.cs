using System.Collections.Generic;
using SpreadsheetCLI.Domain.ValueObjects;

namespace SpreadsheetCLI.Application.DTOs;

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