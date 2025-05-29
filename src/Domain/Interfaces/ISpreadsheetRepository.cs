namespace SpreadsheetCLI.Domain.Interfaces;

using System.Collections.Generic;
using System.Threading.Tasks;
using SpreadsheetCLI.Domain.Entities;
using SpreadsheetCLI.Domain.ValueObjects;

public interface ISpreadsheetRepository
{
    Task<DocumentMetadata> GetDocumentMetadataAsync(string filePath);
    Task<DocumentContext> GetDocumentContextAsync(string filePath);
    Task<IEnumerable<Dictionary<string, object?>>> ReadDataAsync(string filePath, string? sheetName = null, FilterCriteria? filter = null);
    Task<Dictionary<string, ColumnStats>> GetColumnStatisticsAsync(string filePath, string? sheetName = null);
}