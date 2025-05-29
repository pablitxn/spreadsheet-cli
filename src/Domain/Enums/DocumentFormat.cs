namespace SpreadsheetCLI.Domain.Enums;

/// <summary>
/// Document format types that can be detected
/// </summary>
public enum DocumentFormat
{
    Columnar,
    RowBased,
    Nested,
    Matrix,
    Mixed,
    Unknown,
    Excel,
    CSV,
    TSV,
    Other
}