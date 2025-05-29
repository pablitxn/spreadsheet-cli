using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;
using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel.ChatCompletion;
using Moq;
using SpreadsheetCLI.Application.Interfaces;
using SpreadsheetCLI.Application.Interfaces.Spreadsheet;
using SpreadsheetCLI.Application.DTOs;
using SpreadsheetCLI.Domain.ValueObjects;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Services;
using SpreadsheetCLI.Infrastructure.Services;
using Xunit;

namespace SpreadsheetCLI.Tests.Infrastructure.Services;

public class SpreadsheetAnalysisServiceEdgeCaseTests
{
    private readonly Mock<ILogger<SpreadsheetAnalysisService>> _loggerMock;
    private readonly Mock<IChatCompletionService> _chatCompletionMock;
    private readonly Mock<IActivityPublisher> _activityPublisherMock;
    private readonly Mock<FileLoggerService> _fileLoggerMock;
    private readonly SpreadsheetAnalysisService _service;

    public SpreadsheetAnalysisServiceEdgeCaseTests()
    {
        _loggerMock = new Mock<ILogger<SpreadsheetAnalysisService>>();
        _chatCompletionMock = new Mock<IChatCompletionService>();
        _activityPublisherMock = new Mock<IActivityPublisher>();
        _fileLoggerMock = new Mock<FileLoggerService>();
        _service = new SpreadsheetAnalysisService(
            _loggerMock.Object,
            _chatCompletionMock.Object,
            _activityPublisherMock.Object,
            _fileLoggerMock.Object);
    }

    #region Large Dataset Tests

    [Fact]
    public async Task CreateDynamicSpreadsheetAsync_LargeDataset_HandlesEfficiently()
    {
        // Arrange
        var rows = 1000;
        var cols = 50;
        var artifactsFormatted = new List<List<object>>();
        
        // Add headers
        var headers = new List<object>();
        for (int i = 0; i < cols; i++)
        {
            headers.Add($"Column{i}");
        }
        artifactsFormatted.Add(headers);
        
        // Add data rows
        for (int r = 0; r < rows - 1; r++)
        {
            var row = new List<object>();
            for (int c = 0; c < cols; c++)
            {
                row.Add(r * cols + c);
            }
            artifactsFormatted.Add(row);
        }

        var executionPlan = new ExecutionPlanDto
        {
            ArtifactsFormatted = artifactsFormatted
        };

        // Act
        var result = await _service.CreateDynamicSpreadsheetAsync(executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.Equal(rows, result.DataRows);
        Assert.Equal(cols, result.DataColumns);
        Assert.Equal(rows * cols, result.CellAssignments.Count);
    }

    #endregion

    #region Special Character and Unicode Tests

    [Fact]
    public async Task CreateDynamicSpreadsheetAsync_SpecialCharacters_HandlesCorrectly()
    {
        // Arrange
        var executionPlan = new ExecutionPlanDto
        {
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { "Special", "Characters", "Test" },
                new List<object> { "Test@#$%", "Line\nBreak", "Tab\tChar" },
                new List<object> { "Quote\"Test", "Apostrophe'Test", "Backslash\\Test" },
                new List<object> { "Unicode🎉", "Émoji😊", "中文字符" }
            }
        };

        // Act
        var result = await _service.CreateDynamicSpreadsheetAsync(executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.Equal(4, result.DataRows);
        Assert.Equal("Test@#$%", result.CellAssignments[3].AssignedValue);
        Assert.Equal("Line\nBreak", result.CellAssignments[4].AssignedValue);
        Assert.Equal("Unicode🎉", result.CellAssignments[9].AssignedValue);
        Assert.Equal("中文字符", result.CellAssignments[11].AssignedValue);
    }

    #endregion

    #region Numeric Edge Cases

    [Fact]
    public void RowMatchesFilters_NumericEdgeCases_HandlesCorrectly()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        
        // Test various numeric formats
        worksheet.Cells[0, 0].Value = double.MaxValue;
        worksheet.Cells[1, 0].Value = double.MinValue;
        worksheet.Cells[2, 0].Value = 0.0000001;
        worksheet.Cells[3, 0].Value = -0.0000001;
        worksheet.Cells[4, 0].Value = double.NaN;
        worksheet.Cells[5, 0].Value = double.PositiveInfinity;
        worksheet.Cells[6, 0].Value = double.NegativeInfinity;
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Value", 0) };

        // Test 1: MaxValue comparison
        var filters1 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Value", Operator = ">", Value = "1.7E+308" }
        };
        var result1 = _service.RowMatchesFilters(worksheet, 0, headers, filters1);
        Assert.True(result1);

        // Test 2: MinValue comparison
        var filters2 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Value", Operator = "<", Value = "-1.7E+308" }
        };
        var result2 = _service.RowMatchesFilters(worksheet, 1, headers, filters2);
        Assert.True(result2);

        // Test 3: Very small positive number
        var filters3 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Value", Operator = ">", Value = "0" }
        };
        var result3 = _service.RowMatchesFilters(worksheet, 2, headers, filters3);
        Assert.True(result3);

        // Test 4: Very small negative number
        var filters4 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Value", Operator = "<", Value = "0" }
        };
        var result4 = _service.RowMatchesFilters(worksheet, 3, headers, filters4);
        Assert.True(result4);
    }

    #endregion

    #region Formula Edge Cases

    [Fact]
    public async Task ExecuteFormulaAsync_ComplexFormulas_ExecutesCorrectly()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        
        // Setup data for complex formulas
        worksheet.Cells[0, 0].Value = "Product";
        worksheet.Cells[0, 1].Value = "Price";
        worksheet.Cells[0, 2].Value = "Quantity";
        worksheet.Cells[1, 0].Value = "A";
        worksheet.Cells[1, 1].Value = 10.5;
        worksheet.Cells[1, 2].Value = 5;
        worksheet.Cells[2, 0].Value = "B";
        worksheet.Cells[2, 1].Value = 20.0;
        worksheet.Cells[2, 2].Value = 3;
        worksheet.Cells[3, 0].Value = "A";
        worksheet.Cells[3, 1].Value = 15.5;
        worksheet.Cells[3, 2].Value = 2;

        var executionPlan = new ExecutionPlanDto
        {
            Formula = "=SUMPRODUCT(B2:B4,C2:C4)",
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { "Product", "Price", "Quantity" },
                new List<object> { "A", 10.5, 5 },
                new List<object> { "B", 20.0, 3 },
                new List<object> { "A", 15.5, 2 }
            }
        };

        // Act
        var result = await _service.ExecuteFormulaAsync(workbook, executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.True(result.Success);
        // 10.5*5 + 20*3 + 15.5*2 = 52.5 + 60 + 31 = 143.5
        Assert.Equal(143.5, result.Value);
    }

    [Fact]
    public async Task ExecuteFormulaAsync_NestedFormulas_ExecutesCorrectly()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        
        worksheet.Cells[0, 0].Value = 10;
        worksheet.Cells[1, 0].Value = 20;
        worksheet.Cells[2, 0].Value = 30;
        worksheet.Cells[3, 0].Value = 40;

        var executionPlan = new ExecutionPlanDto
        {
            Formula = "=IF(SUM(A1:A4)>50,AVERAGE(A1:A4),MAX(A1:A4))",
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { 10 },
                new List<object> { 20 },
                new List<object> { 30 },
                new List<object> { 40 }
            }
        };

        // Act
        var result = await _service.ExecuteFormulaAsync(workbook, executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.True(result.Success);
        // SUM = 100 > 50, so AVERAGE = 25
        Assert.Equal(25.0, result.Value);
    }

    #endregion

    #region Date Filter Tests

    [Fact]
    public void RowMatchesFilters_DateFilters_HandlesCorrectly()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        
        // Set various date formats
        worksheet.Cells[0, 0].Value = new DateTime(2024, 1, 15);
        worksheet.Cells[1, 0].Value = DateTime.Parse("2024-06-30");
        worksheet.Cells[2, 0].Value = 45292; // OLE Automation date for 2024-01-01
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Date", 0) };

        // Test date greater than
        var filters1 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Date", Operator = "date>", Value = "2024-01-01" }
        };
        var result1 = _service.RowMatchesFilters(worksheet, 0, headers, filters1);
        Assert.True(result1);

        // Test date less than
        var filters2 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Date", Operator = "date<", Value = "2024-12-31" }
        };
        var result2 = _service.RowMatchesFilters(worksheet, 1, headers, filters2);
        Assert.True(result2);
    }

    #endregion

    #region Empty and Null Handling

    [Fact]
    public async Task CreateDynamicSpreadsheetAsync_EmptyRows_HandlesGracefully()
    {
        // Arrange
        var executionPlan = new ExecutionPlanDto
        {
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { "Col1", "Col2", "Col3" },
                new List<object> { "", null, "" },
                new List<object> { null, null, null },
                new List<object> { "Data", "", null }
            }
        };

        // Act
        var result = await _service.CreateDynamicSpreadsheetAsync(executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.Equal(4, result.DataRows);
        Assert.Equal(3, result.DataColumns);
        
        // Check null handling
        Assert.Equal("", result.CellAssignments[3].AssignedValue);
        Assert.Null(result.CellAssignments[4].AssignedValue);
        Assert.Equal("", result.CellAssignments[5].AssignedValue);
        Assert.Null(result.CellAssignments[6].AssignedValue);
    }

    #endregion

    #region Accounting Format Tests

    [Fact]
    public void RowMatchesFilters_AccountingFormat_ParsesCorrectly()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        
        // Test accounting formats
        worksheet.Cells[0, 0].Value = "$1,234.56";
        worksheet.Cells[1, 0].Value = "($500.00)"; // Negative in parentheses
        worksheet.Cells[2, 0].Value = "$0.00";
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Amount", 0) };

        // Test positive amount
        var filters1 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Amount", Operator = ">", Value = "1000" }
        };
        var result1 = _service.RowMatchesFilters(worksheet, 0, headers, filters1);
        Assert.True(result1);

        // Test negative amount
        var filters2 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Amount", Operator = "<", Value = "0" }
        };
        var result2 = _service.RowMatchesFilters(worksheet, 1, headers, filters2);
        Assert.True(result2);

        // Test zero
        var filters3 = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Amount", Operator = "equals", Value = "0" }
        };
        var result3 = _service.RowMatchesFilters(worksheet, 2, headers, filters3);
        Assert.True(result3);
    }

    #endregion

    #region Performance Tests

    [Fact]
    public void RowMatchesFilters_ManyFilters_PerformsEfficiently()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        
        // Setup data with many columns
        var headers = new List<HeaderInfo>();
        var filters = new List<FilterCriteria>();
        
        for (int i = 0; i < 100; i++)
        {
            var colName = $"Col{i}";
            headers.Add(new HeaderInfo(colName, 0));
            worksheet.Cells[1, i].Value = i;
            
            // Add filter for every 10th column
            if (i % 10 == 0)
            {
                filters.Add(new FilterCriteria 
                { 
                    Column = colName, 
                    Operator = ">=", 
                    Value = "0" 
                });
            }
        }

        // Act
        var startTime = DateTime.UtcNow;
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);
        var duration = DateTime.UtcNow - startTime;

        // Assert
        Assert.True(result);
        Assert.True(duration.TotalMilliseconds < 100, $"Filter evaluation took {duration.TotalMilliseconds}ms");
    }

    #endregion

    #region Case Sensitivity Tests

    [Fact]
    public void RowMatchesFilters_CaseInsensitive_WorksCorrectly()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = "John DOE";
        
        var headers = new List<HeaderInfo> { new HeaderInfo("name", 0) }; // lowercase header
        
        // Test with different case in filter
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "NAME", Operator = "contains", Value = "john" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.True(result); // Should be case-insensitive
    }

    #endregion
}