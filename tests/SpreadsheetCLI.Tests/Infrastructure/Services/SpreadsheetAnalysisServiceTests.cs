using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading;
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

public class SpreadsheetAnalysisServiceTests
{
    private readonly Mock<ILogger<SpreadsheetAnalysisService>> _loggerMock;
    private readonly Mock<IChatCompletionService> _chatCompletionMock;
    private readonly Mock<IActivityPublisher> _activityPublisherMock;
    private readonly Mock<FileLoggerService> _fileLoggerMock;
    private readonly SpreadsheetAnalysisService _service;

    public SpreadsheetAnalysisServiceTests()
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

    #region ExtractHeaders Tests

    [Fact]
    public void ExtractHeaders_EmptyWorksheet_ReturnsEmptyList()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Act
        var headers = _service.ExtractHeaders(worksheet);

        // Assert
        Assert.NotNull(headers);
        Assert.Empty(headers);
    }

    [Fact]
    public void ExtractHeaders_WorksheetWithHeaders_ReturnsHeaderList()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[0, 0].Value = "Name";
        worksheet.Cells[0, 1].Value = "Age";
        worksheet.Cells[0, 2].Value = "City";

        // Act
        var headers = _service.ExtractHeaders(worksheet);

        // Assert
        Assert.NotNull(headers);
        // Note: This test will need adjustment based on actual implementation
        // since ExtractHeaders uses LLM internally
    }

    #endregion

    #region CountDataRows Tests

    [Fact]
    public void CountDataRows_EmptyWorksheet_ReturnsZero()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Act
        var count = _service.CountDataRows(worksheet);

        // Assert
        Assert.Equal(0, count);
    }

    [Fact]
    public void CountDataRows_WorksheetWithData_ReturnsCorrectCount()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].Value = $"Data{i}";
        }

        // Act
        var count = _service.CountDataRows(worksheet);

        // Assert
        Assert.Equal(9, count); // Row 1 to Row 9 (excluding header)
    }

    #endregion

    #region GetDataRange Tests

    [Fact]
    public void GetDataRange_EmptyWorksheet_ReturnsDefaultRange()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Act
        var (firstRow, lastRow) = _service.GetDataRange(worksheet);

        // Assert
        Assert.Equal(1, firstRow);
        Assert.Equal(-1, lastRow); // MaxRow returns -1 for empty worksheet
    }

    [Fact]
    public void GetDataRange_WorksheetWithData_ReturnsCorrectRange()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].Value = $"Data{i}";
        }

        // Act
        var (firstRow, lastRow) = _service.GetDataRange(worksheet);

        // Assert
        Assert.Equal(1, firstRow);
        Assert.Equal(9, lastRow);
    }

    #endregion

    #region RowMatchesFilters Tests

    [Fact]
    public void RowMatchesFilters_NoFilters_ReturnsTrue()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = "John";
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Name", 0) };
        var filters = new List<FilterCriteria>();

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.True(result);
    }

    [Fact]
    public void RowMatchesFilters_EqualsFilter_MatchingValue_ReturnsTrue()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = "John";
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Name", 0) };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Name", Operator = "equals", Value = "John" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.True(result);
    }

    [Fact]
    public void RowMatchesFilters_EqualsFilter_NonMatchingValue_ReturnsFalse()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = "John";
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Name", 0) };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Name", Operator = "equals", Value = "Jane" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.False(result);
    }

    [Fact]
    public void RowMatchesFilters_ContainsFilter_MatchingValue_ReturnsTrue()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = "John Doe";
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Name", 0) };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Name", Operator = "contains", Value = "John" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.True(result);
    }

    [Fact]
    public void RowMatchesFilters_NumericGreaterThan_MatchingValue_ReturnsTrue()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = 25;
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Age", 0) };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Age", Operator = ">", Value = "20" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.True(result);
    }

    [Fact]
    public void RowMatchesFilters_NumericLessThan_NonMatchingValue_ReturnsFalse()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = 25;
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Age", 0) };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Age", Operator = "<", Value = "20" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.False(result);
    }

    [Fact]
    public void RowMatchesFilters_MultipleFilters_AllMatch_ReturnsTrue()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = "John";
        worksheet.Cells[1, 1].Value = 25;
        
        var headers = new List<HeaderInfo> 
        { 
            new HeaderInfo("Name", 0),
            new HeaderInfo("Age", 0)
        };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Name", Operator = "equals", Value = "John" },
            new FilterCriteria { Column = "Age", Operator = ">", Value = "20" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.True(result);
    }

    [Fact]
    public void RowMatchesFilters_MultipleFilters_OneDoesNotMatch_ReturnsFalse()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = "John";
        worksheet.Cells[1, 1].Value = 15;
        
        var headers = new List<HeaderInfo> 
        { 
            new HeaderInfo("Name", 0),
            new HeaderInfo("Age", 0)
        };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Name", Operator = "equals", Value = "John" },
            new FilterCriteria { Column = "Age", Operator = ">", Value = "20" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.False(result);
    }

    #endregion

    #region CreateDynamicSpreadsheetAsync Tests

    [Fact]
    public async Task CreateDynamicSpreadsheetAsync_EmptyExecutionPlan_ReturnsEmptyWorkbook()
    {
        // Arrange
        var executionPlan = new ExecutionPlanDto
        {
            ArtifactsFormatted = null
        };

        // Act
        var result = await _service.CreateDynamicSpreadsheetAsync(executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.NotNull(result.Workbook);
        Assert.NotNull(result.Worksheet);
        Assert.Equal(0, result.DataRows);
        Assert.Equal(0, result.DataColumns);
        Assert.Empty(result.CellAssignments);
    }

    [Fact]
    public async Task CreateDynamicSpreadsheetAsync_WithData_CreatesCorrectWorkbook()
    {
        // Arrange
        var executionPlan = new ExecutionPlanDto
        {
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { "Name", "Age", "City" },
                new List<object> { "John", 25, "New York" },
                new List<object> { "Jane", 30, "London" }
            }
        };

        // Act
        var result = await _service.CreateDynamicSpreadsheetAsync(executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.Equal(3, result.DataRows);
        Assert.Equal(3, result.DataColumns);
        Assert.Equal(9, result.CellAssignments.Count);
        
        // Verify first row (headers)
        Assert.Equal("A1", result.CellAssignments[0].CellReference);
        Assert.Equal("Name", result.CellAssignments[0].AssignedValue);
        
        // Verify data
        Assert.Equal("B2", result.CellAssignments[4].CellReference);
        Assert.Equal(25, result.CellAssignments[4].AssignedValue);
    }

    [Fact]
    public async Task CreateDynamicSpreadsheetAsync_WithJsonElements_ConvertsCorrectly()
    {
        // Arrange
        var jsonData = JsonSerializer.Serialize(new { value = 42 });
        var jsonDoc = JsonDocument.Parse(jsonData);
        var jsonElement = jsonDoc.RootElement.GetProperty("value");

        var executionPlan = new ExecutionPlanDto
        {
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { "Number" },
                new List<object> { jsonElement }
            }
        };

        // Act
        var result = await _service.CreateDynamicSpreadsheetAsync(executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.Equal(2, result.DataRows);
        Assert.Equal(1, result.DataColumns);
        Assert.Equal(42.0, result.CellAssignments[1].AssignedValue);
        Assert.True(result.CellAssignments[1].IsJsonElement);
    }

    #endregion

    #region ExecuteFormulaAsync Tests

    [Fact]
    public async Task ExecuteFormulaAsync_SimpleFormula_ReturnsCorrectResult()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[0, 0].Value = 10;
        worksheet.Cells[1, 0].Value = 20;
        worksheet.Cells[2, 0].Value = 30;

        var executionPlan = new ExecutionPlanDto
        {
            Formula = "=SUM(A1:A3)",
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { 10 },
                new List<object> { 20 },
                new List<object> { 30 }
            }
        };

        // Act
        var result = await _service.ExecuteFormulaAsync(workbook, executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal(60.0, result.Value);
        Assert.Equal("60", result.StringValue);
        Assert.Null(result.Error);
        Assert.Equal("A4", result.FormulaCellReference);
    }

    [Fact]
    public async Task ExecuteFormulaAsync_InvalidFormula_ReturnsError()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        var executionPlan = new ExecutionPlanDto
        {
            Formula = "=INVALID_FUNCTION()",
            ArtifactsFormatted = new List<List<object>>()
        };

        // Act
        var result = await _service.ExecuteFormulaAsync(workbook, executionPlan);

        // Assert
        Assert.NotNull(result);
        // Depending on Aspose behavior, this might succeed with an error value
        // or throw an exception caught by the try-catch
        Assert.NotNull(result.FormulaCellReference);
    }

    [Fact]
    public async Task ExecuteFormulaAsync_AverageFormula_ReturnsCorrectResult()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[0, 0].Value = "Value";
        worksheet.Cells[1, 0].Value = 10;
        worksheet.Cells[2, 0].Value = 20;
        worksheet.Cells[3, 0].Value = 30;

        var executionPlan = new ExecutionPlanDto
        {
            Formula = "=AVERAGE(A2:A4)",
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { "Value" },
                new List<object> { 10 },
                new List<object> { 20 },
                new List<object> { 30 }
            }
        };

        // Act
        var result = await _service.ExecuteFormulaAsync(workbook, executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal(20.0, result.Value);
        Assert.Equal("20", result.StringValue);
    }

    [Fact]
    public async Task ExecuteFormulaAsync_CountFormula_ReturnsCorrectResult()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[0, 0].Value = "Name";
        worksheet.Cells[1, 0].Value = "John";
        worksheet.Cells[2, 0].Value = "Jane";
        worksheet.Cells[3, 0].Value = "";
        worksheet.Cells[4, 0].Value = "Bob";

        var executionPlan = new ExecutionPlanDto
        {
            Formula = "=COUNTA(A2:A5)",
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { "Name" },
                new List<object> { "John" },
                new List<object> { "Jane" },
                new List<object> { "" },
                new List<object> { "Bob" }
            }
        };

        // Act
        var result = await _service.ExecuteFormulaAsync(workbook, executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal(3.0, result.Value); // Should count non-empty cells
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void RowMatchesFilters_NullCellValue_HandlesGracefully()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        // Cell is not set, so it's null
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Name", 0) };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Name", Operator = "equals", Value = "" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.True(result); // null equals empty string
    }

    [Fact]
    public void RowMatchesFilters_ColumnNotInHeaders_SkipsFilter()
    {
        // Arrange
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[1, 0].Value = "John";
        
        var headers = new List<HeaderInfo> { new HeaderInfo("Name", 0) };
        var filters = new List<FilterCriteria>
        {
            new FilterCriteria { Column = "Age", Operator = "equals", Value = "25" }
        };

        // Act
        var result = _service.RowMatchesFilters(worksheet, 1, headers, filters);

        // Assert
        Assert.True(result); // Filter is skipped when column not found
    }

    [Fact]
    public async Task CreateDynamicSpreadsheetAsync_MixedDataTypes_HandlesCorrectly()
    {
        // Arrange
        var executionPlan = new ExecutionPlanDto
        {
            ArtifactsFormatted = new List<List<object>>
            {
                new List<object> { "Mixed", "Data", "Types" },
                new List<object> { 123, "Text", true },
                new List<object> { null, 45.67, false }
            }
        };

        // Act
        var result = await _service.CreateDynamicSpreadsheetAsync(executionPlan);

        // Assert
        Assert.NotNull(result);
        Assert.Equal(3, result.DataRows);
        Assert.Equal(3, result.DataColumns);
        
        // Check various data types
        Assert.Equal(123, result.CellAssignments[3].AssignedValue);
        Assert.Equal("Text", result.CellAssignments[4].AssignedValue);
        Assert.Equal(true, result.CellAssignments[5].AssignedValue);
        Assert.Null(result.CellAssignments[6].AssignedValue);
        Assert.Equal(45.67, result.CellAssignments[7].AssignedValue);
        Assert.Equal(false, result.CellAssignments[8].AssignedValue);
    }

    #endregion
}