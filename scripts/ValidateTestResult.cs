using System;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using SpreadsheetCLI.Application.DTOs;
using SpreadsheetCLI.Application.Interfaces;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Services;

public class Program
{
    public static async Task<int> Main(string[] args)
    {
        if (args.Length < 3)
        {
            Console.Error.WriteLine("Usage: ValidateTestResult <question> <expected_answer> <actual_output>");
            return 1;
        }

        var question = args[0];
        var expectedAnswer = args[1];
        var actualOutput = args[2];

        try
        {
            // Build configuration
            var configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: true)
                .AddEnvironmentVariables()
                .Build();

            // Create service collection
            var services = new ServiceCollection();
            
            // Add logging
            services.AddLogging(builder =>
            {
                builder.SetMinimumLevel(LogLevel.Warning);
                builder.AddConsole();
            });

            // Get OpenAI configuration
            var openAiApiKey = configuration["OpenAI:ApiKey"] ?? 
                               configuration["SemanticKernel:OpenAI:ApiKey"] ?? 
                               Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            
            if (string.IsNullOrEmpty(openAiApiKey))
            {
                Console.Error.WriteLine("Error: OpenAI API key not found in configuration or environment variables");
                return 1;
            }

            // Add Semantic Kernel
            services.AddSingleton<IKernelBuilder>(sp =>
            {
                var builder = Kernel.CreateBuilder();
                builder.Services.AddOpenAIChatCompletion(
                    modelId: "gpt-4o-mini",
                    apiKey: openAiApiKey);
                return builder;
            });

            services.AddSingleton<IKernel>(sp =>
            {
                var builder = sp.GetRequiredService<IKernelBuilder>();
                return builder.Build();
            });

            services.AddSingleton<IChatCompletionService>(sp =>
            {
                var kernel = sp.GetRequiredService<IKernel>();
                return kernel.GetRequiredService<IChatCompletionService>();
            });

            // Add validation service
            services.AddSingleton<ITestResultValidationService, TestResultValidationService>();

            // Build service provider
            var serviceProvider = services.BuildServiceProvider();

            // Get validation service
            var validationService = serviceProvider.GetRequiredService<ITestResultValidationService>();

            // Create validation request
            var request = new TestValidationRequest
            {
                Question = question,
                ExpectedAnswer = expectedAnswer,
                ActualOutput = actualOutput
            };

            // Perform validation
            var result = await validationService.ValidateTestResultAsync(request);

            // Output result as JSON
            var jsonResult = JsonSerializer.Serialize(result, new JsonSerializerOptions
            {
                WriteIndented = false // Single line for easier parsing in bash
            });

            Console.WriteLine(jsonResult);

            // Return 0 if test passed, 1 if failed
            return result.IsCorrect ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            
            // Output error result as JSON
            var errorResult = new TestValidationResult
            {
                IsCorrect = false,
                ExtractedAnswer = "Error",
                Explanation = ex.Message,
                Confidence = 0,
                AnswerLocation = "N/A"
            };
            
            Console.WriteLine(JsonSerializer.Serialize(errorResult));
            return 1;
        }
    }
}