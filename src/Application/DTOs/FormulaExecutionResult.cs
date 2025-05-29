using System.Collections.Generic;

namespace SpreadsheetCLI.Application.DTOs;

/// <summary>
/// Result of formula execution
/// </summary>
public sealed class FormulaExecutionResult
{
    public bool Success { get; set; }
    public object? Value { get; set; }
    public string StringValue { get; set; } = "";
    public string? Error { get; set; }
    public string FormulaCellReference { get; set; } = "";
    public Dictionary<string, object> DebugInfo { get; set; } = new();
}