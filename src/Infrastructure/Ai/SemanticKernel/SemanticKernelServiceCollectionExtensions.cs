using System;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Http;
using Microsoft.SemanticKernel;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Plugins;

namespace Infrastructure.Ai.SemanticKernel;

/// <summary>
/// Registers Semantic Kernel services for the CLI
/// </summary>
public static class SemanticKernelServiceCollectionExtensions
{
    public static IServiceCollection AddSemanticKernel(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        // Get API key from configuration or environment
        var apiKey = configuration["OpenAI:ApiKey"] 
            ?? Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            
        if (string.IsNullOrWhiteSpace(apiKey))
            throw new InvalidOperationException(
                "OpenAI API key missing. Set OpenAI:ApiKey in appsettings.json or the OPENAI_API_KEY env-var.");

        var modelId = configuration["OpenAI:Model"] ?? "gpt-4o";

        // Add OpenAI chat completion
        services.AddOpenAIChatCompletion(modelId, apiKey);

        // Add baseline infrastructure
        services.AddMemoryCache();
        services.AddDistributedMemoryCache();
        services.AddHttpClient();

        // Register the Kernel
        services.AddScoped<Kernel>(sp =>
        {
            var kernel = new Kernel(sp);
            kernel.ImportPluginFromType<SpreadsheetPlugin>("excel");
            return kernel;
        });

        return services;
    }
    
    public static IServiceCollection AddSemanticKernelPlugins(this IServiceCollection services)
    {
        return services
            .AddTransient<SpreadsheetPlugin>();
    }
}