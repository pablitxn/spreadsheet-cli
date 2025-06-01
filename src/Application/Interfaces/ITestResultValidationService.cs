using System.Threading;
using System.Threading.Tasks;
using SpreadsheetCLI.Application.DTOs;

namespace SpreadsheetCLI.Application.Interfaces;

/// <summary>
/// Service for validating test results using LLM
/// </summary>
public interface ITestResultValidationService
{
    /// <summary>
    /// Validates a test result by comparing actual output with expected answer
    /// </summary>
    /// <param name="request">The validation request containing question, expected answer, and actual output</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Validation result with extracted answer and pass/fail status</returns>
    Task<TestValidationResult> ValidateTestResultAsync(
        TestValidationRequest request, 
        CancellationToken cancellationToken = default);
}