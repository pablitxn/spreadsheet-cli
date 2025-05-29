using Microsoft.Extensions.DependencyInjection;
using SpreadsheetCLI.Application.Interfaces.Spreadsheet;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Services;

namespace SpreadsheetCLI.Infrastructure.Ai.SemanticKernel;

/// <summary>
/// Extension methods for registering spreadsheet services
/// </summary>
public static class SpreadsheetServicesExtensions
{
    /// <summary>
    /// Adds spreadsheet analysis and execution services to the DI container
    /// </summary>
    public static IServiceCollection AddSpreadsheetServices(this IServiceCollection services)
    {
        // Register core services
        services.AddScoped<ISpreadsheetAnalysisService, SpreadsheetAnalysisService>();

        return services;
    }
}