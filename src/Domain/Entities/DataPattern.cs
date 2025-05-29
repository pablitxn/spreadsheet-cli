namespace SpreadsheetCLI.Domain.Entities;

/// <summary>
/// Detected data pattern
/// </summary>
public sealed class DataPattern
{
    public string PatternType { get; set; } = "";
    public string Description { get; set; } = "";
    public string Location { get; set; } = "";
    public double Confidence { get; set; }
}