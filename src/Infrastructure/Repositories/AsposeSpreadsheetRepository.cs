namespace SpreadsheetCLI.Infrastructure.Repositories;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;
using Microsoft.Extensions.Logging;
using SpreadsheetCLI.Domain.Entities;
using SpreadsheetCLI.Domain.Enums;
using SpreadsheetCLI.Domain.Interfaces;
using SpreadsheetCLI.Domain.ValueObjects;
using System.Data;

public class AsposeSpreadsheetRepository : ISpreadsheetRepository
{
    private readonly ILogger<AsposeSpreadsheetRepository> _logger;

    public AsposeSpreadsheetRepository(ILogger<AsposeSpreadsheetRepository> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public async Task<DocumentMetadata> GetDocumentMetadataAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            using var workbook = new Workbook(filePath);
            var fileInfo = new FileInfo(filePath);
            
            return new DocumentMetadata
            {
                FilePath = filePath,
                FileName = Path.GetFileName(filePath),
                Format = GetDocumentFormat(fileInfo.Extension),
                FileSize = fileInfo.Length,
                SheetCount = workbook.Worksheets.Count,
                CreatedDate = fileInfo.CreationTime,
                ModifiedDate = fileInfo.LastWriteTime,
                Properties = ExtractProperties(workbook)
            };
        });
    }

    public async Task<DocumentContext> GetDocumentContextAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            using var workbook = new Workbook(filePath);
            var sheets = new Dictionary<string, SheetInfo>();
            var relationships = new List<Relationship>();
            
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                var sheetInfo = ExtractSheetInfo(worksheet);
                sheets[worksheet.Name] = sheetInfo;
                relationships.AddRange(ExtractRelationships(worksheet, workbook));
            }
            
            return new DocumentContext
            {
                Metadata = GetDocumentMetadataAsync(filePath).Result,
                Sheets = sheets,
                Relationships = relationships,
                DataPatterns = ExtractDataPatterns(workbook)
            };
        });
    }

    public async Task<IEnumerable<Dictionary<string, object?>>> ReadDataAsync(string filePath, string? sheetName = null, FilterCriteria? filter = null)
    {
        return await Task.Run(() =>
        {
            using var workbook = new Workbook(filePath);
            var worksheet = sheetName != null 
                ? workbook.Worksheets[sheetName] 
                : workbook.Worksheets[0];
                
            if (worksheet == null)
                throw new ArgumentException($"Sheet '{sheetName}' not found");
                
            var data = new List<Dictionary<string, object?>>();
            var maxRow = worksheet.Cells.MaxDataRow;
            var maxCol = worksheet.Cells.MaxDataColumn;
            
            // Get headers
            var headers = new List<string>();
            for (int col = 0; col <= maxCol; col++)
            {
                var cell = worksheet.Cells[0, col];
                headers.Add(cell.StringValue ?? $"Column{col + 1}");
            }
            
            // Read data
            for (int row = 1; row <= maxRow; row++)
            {
                var rowData = new Dictionary<string, object?>();
                bool includeRow = true;
                
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var value = GetCellValue(cell);
                    rowData[headers[col]] = value;
                    
                    // Apply filter if provided
                    if (filter != null && filter.Column == headers[col])
                    {
                        includeRow = EvaluateFilter(value, filter);
                    }
                }
                
                if (includeRow)
                    data.Add(rowData);
            }
            
            return data;
        });
    }

    public async Task<Dictionary<string, ColumnStats>> GetColumnStatisticsAsync(string filePath, string? sheetName = null)
    {
        return await Task.Run(() =>
        {
            using var workbook = new Workbook(filePath);
            var worksheet = sheetName != null 
                ? workbook.Worksheets[sheetName] 
                : workbook.Worksheets[0];
                
            if (worksheet == null)
                throw new ArgumentException($"Sheet '{sheetName}' not found");
                
            var stats = new Dictionary<string, ColumnStats>();
            var maxRow = worksheet.Cells.MaxDataRow;
            var maxCol = worksheet.Cells.MaxDataColumn;
            
            // Get headers
            var headers = new List<string>();
            for (int col = 0; col <= maxCol; col++)
            {
                var cell = worksheet.Cells[0, col];
                headers.Add(cell.StringValue ?? $"Column{col + 1}");
            }
            
            // Calculate stats for each column
            for (int col = 0; col <= maxCol; col++)
            {
                var columnStats = CalculateColumnStats(worksheet, col, 1, maxRow);
                stats[headers[col]] = columnStats;
            }
            
            return stats;
        });
    }

    private DocumentFormat GetDocumentFormat(string extension)
    {
        return extension.ToLower() switch
        {
            ".xlsx" => DocumentFormat.Excel,
            ".xls" => DocumentFormat.Excel,
            ".csv" => DocumentFormat.CSV,
            ".tsv" => DocumentFormat.TSV,
            _ => DocumentFormat.Other
        };
    }

    private Dictionary<string, string> ExtractProperties(Workbook workbook)
    {
        var props = new Dictionary<string, string>();
        var docProps = workbook.BuiltInDocumentProperties;
        
        if (!string.IsNullOrEmpty(docProps.Author))
            props["Author"] = docProps.Author;
        if (!string.IsNullOrEmpty(docProps.Title))
            props["Title"] = docProps.Title;
        if (!string.IsNullOrEmpty(docProps.Subject))
            props["Subject"] = docProps.Subject;
            
        return props;
    }

    private SheetInfo ExtractSheetInfo(Worksheet worksheet)
    {
        var maxRow = worksheet.Cells.MaxDataRow;
        var maxCol = worksheet.Cells.MaxDataColumn;
        
        var headers = new List<HeaderInfo>();
        var columnTypes = new Dictionary<string, ColumnType>();
        for (int col = 0; col <= maxCol; col++)
        {
            var cell = worksheet.Cells[0, col];
            var headerName = cell.StringValue ?? $"Column{col + 1}";
            var columnType = DetermineColumnType(worksheet, col, 1, Math.Min(10, maxRow));
            
            headers.Add(new HeaderInfo(headerName, 0)); // Row index is 0 for headers
            columnTypes[headerName] = columnType;
        }
        
        return new SheetInfo
        {
            Name = worksheet.Name,
            RowCount = maxRow + 1,
            ColumnCount = maxCol + 1,
            Headers = headers,
            ColumnTypes = columnTypes,
            FormulaCells = GetFormulaCells(worksheet),
            IsHidden = !worksheet.IsVisible
        };
    }

    private ColumnType DetermineColumnType(Worksheet worksheet, int col, int startRow, int endRow)
    {
        var types = new Dictionary<ColumnType, int>();
        
        for (int row = startRow; row <= endRow; row++)
        {
            var cell = worksheet.Cells[row, col];
            if (cell.Value == null) continue;
            
            var type = cell.Type switch
            {
                CellValueType.IsNumeric => ColumnType.Numeric,
                CellValueType.IsDateTime => ColumnType.Date,
                CellValueType.IsBool => ColumnType.Boolean,
                _ => ColumnType.Text
            };
            
            types[type] = types.GetValueOrDefault(type) + 1;
        }
        
        return types.Count > 0 
            ? types.OrderByDescending(t => t.Value).First().Key 
            : ColumnType.Text;
    }

    private List<string> GetFormulaCells(Worksheet worksheet)
    {
        var formulaCells = new List<string>();
        var cells = worksheet.Cells;
        cells.CreateRange(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1);
        
        foreach (Cell cell in cells)
        {
            if (!string.IsNullOrEmpty(cell.Formula))
                formulaCells.Add(cell.Name);
        }
        
        return formulaCells;
    }

    private List<Relationship> ExtractRelationships(Worksheet worksheet, Workbook workbook)
    {
        var relationships = new List<Relationship>();
        
        // This is a simplified implementation
        // In a real scenario, you would analyze formulas, data validation, etc.
        
        return relationships;
    }

    private List<DataPattern> ExtractDataPatterns(Workbook workbook)
    {
        var patterns = new List<DataPattern>();
        
        // This is a simplified implementation
        // In a real scenario, you would analyze data patterns across sheets
        
        return patterns;
    }

    private object? GetCellValue(Cell cell)
    {
        if (cell.Value == null) return null;
        
        return cell.Type switch
        {
            CellValueType.IsNumeric => cell.DoubleValue,
            CellValueType.IsDateTime => cell.DateTimeValue,
            CellValueType.IsBool => cell.BoolValue,
            _ => cell.StringValue
        };
    }

    private bool EvaluateFilter(object? value, FilterCriteria filter)
    {
        if (value == null) return false;
        
        var valueStr = value.ToString() ?? "";
        var filterValue = filter.Value ?? "";
        
        return filter.Operator switch
        {
            "=" => valueStr.Equals(filterValue, StringComparison.OrdinalIgnoreCase),
            "!=" => !valueStr.Equals(filterValue, StringComparison.OrdinalIgnoreCase),
            "contains" => valueStr.Contains(filterValue, StringComparison.OrdinalIgnoreCase),
            ">" => CompareValues(value, filter.Value) > 0,
            "<" => CompareValues(value, filter.Value) < 0,
            ">=" => CompareValues(value, filter.Value) >= 0,
            "<=" => CompareValues(value, filter.Value) <= 0,
            _ => true
        };
    }

    private int CompareValues(object? value1, object? value2)
    {
        if (value1 == null || value2 == null) return 0;
        
        if (value1 is IComparable comparable1 && value2 is IComparable comparable2)
        {
            try
            {
                return comparable1.CompareTo(Convert.ChangeType(comparable2, comparable1.GetType()));
            }
            catch
            {
                return string.Compare(value1.ToString(), value2.ToString(), StringComparison.OrdinalIgnoreCase);
            }
        }
        
        return string.Compare(value1.ToString(), value2.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    private ColumnStats CalculateColumnStats(Worksheet worksheet, int col, int startRow, int endRow)
    {
        var stats = new ColumnStats
        {
            ColumnName = worksheet.Cells[0, col].StringValue ?? $"Column{col + 1}",
            TotalCount = endRow - startRow + 1
        };
        
        var numericValues = new List<double>();
        var uniqueValues = new HashSet<string>();
        int nullCount = 0;
        
        for (int row = startRow; row <= endRow; row++)
        {
            var cell = worksheet.Cells[row, col];
            var value = GetCellValue(cell);
            
            if (value == null)
            {
                nullCount++;
            }
            else
            {
                uniqueValues.Add(value.ToString() ?? "");
                
                if (value is double numValue)
                {
                    numericValues.Add(numValue);
                }
                else if (double.TryParse(value.ToString(), out double parsed))
                {
                    numericValues.Add(parsed);
                }
            }
        }
        
        stats.NonNullCount = stats.TotalCount - nullCount;
        stats.UniqueCount = uniqueValues.Count;
        stats.NullCount = nullCount;
        
        if (numericValues.Count > 0)
        {
            stats.Sum = numericValues.Sum();
            stats.Average = numericValues.Average();
            stats.Min = numericValues.Min();
            stats.Max = numericValues.Max();
        }
        
        return stats;
    }
}