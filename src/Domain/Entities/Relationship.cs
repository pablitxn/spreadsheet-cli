namespace SpreadsheetCLI.Domain.Entities;

/// <summary>
/// Cross-sheet relationship
/// </summary>
public sealed class Relationship
{
    public string SourceSheet { get; set; } = "";
    public string SourceColumn { get; set; } = "";
    public string TargetSheet { get; set; } = "";
    public string TargetColumn { get; set; } = "";
    public string RelationType { get; set; } = "";
}