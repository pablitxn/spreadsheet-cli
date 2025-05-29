using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel;
using SpreadsheetCLI.Core.Application.Interfaces;
using SpreadsheetCLI.Core.Application.Interfaces.Spreadsheet;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Plugins;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Services;
using SpreadsheetCLI.Infrastructure.Mocks;

namespace SpreadsheetCLI;

public class Program
{
    static async Task Main(string[] args)
    {
        var host = CreateHostBuilder(args).Build();

        var logger = host.Services.GetRequiredService<ILogger<Program>>();
        logger.LogInformation("SpreadsheetCLI Started");

        if (args.Length == 0)
        {
            await RunInteractiveMode(host);
        }
        else
        {
            await RunCommandMode(host, args);
        }
    }

    static IHostBuilder CreateHostBuilder(string[] args) =>
        Host.CreateDefaultBuilder(args)
            .ConfigureServices((context, services) =>
            {
                // Add logging
                services.AddLogging(builder =>
                {
                    builder.ClearProviders();
                    builder.AddConsole();
                    builder.SetMinimumLevel(LogLevel.Warning);
                });

                // Add caching
                services.AddMemoryCache();
                services.AddSingleton<IDistributedCache, MemoryDistributedCache>();

                // Add mock services
                services.AddSingleton<IFileStorageService, LocalFileStorageService>();
                services.AddSingleton<IActivityPublisher, FileAndConsoleActivityPublisher>();

                // Add spreadsheet services
                services.AddSingleton<ISpreadsheetAnalysisService, SpreadsheetAnalysisService>();

                // Add Semantic Kernel
                var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    apiKey = context.Configuration["OpenAI:ApiKey"];
                }
                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    throw new InvalidOperationException("OpenAI API key not configured. Set OPENAI_API_KEY environment variable or configure in appsettings.json");
                }

                // Add OpenAI Chat Completion
                services.AddSingleton<Kernel>(sp =>
                {
                    var builder = Kernel.CreateBuilder();
                    builder.AddOpenAIChatCompletion("gpt-4o", apiKey);
                    return builder.Build();
                });

                // Add chat completion service for the analysis services
                services.AddSingleton<Microsoft.SemanticKernel.ChatCompletion.IChatCompletionService>(sp =>
                {
                    var kernel = sp.GetRequiredService<Kernel>();
                    return kernel.GetRequiredService<Microsoft.SemanticKernel.ChatCompletion.IChatCompletionService>();
                });

                // Add SpreadsheetPlugin
                services.AddSingleton<SpreadsheetPlugin>();
            });

    static async Task RunInteractiveMode(IHost host)
    {
        var plugin = host.Services.GetRequiredService<SpreadsheetPlugin>();
        var logger = host.Services.GetRequiredService<ILogger<Program>>();
        var activityPublisher = host.Services.GetRequiredService<IActivityPublisher>() as FileAndConsoleActivityPublisher;

        Console.WriteLine("=== SpreadsheetCLI Interactive Mode ===");
        
        if (activityPublisher != null)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"📝 Audit log file: {activityPublisher.GetLogFilePath()}");
            Console.ResetColor();
            Console.WriteLine();
        }
        
        Console.WriteLine("Enter the path to your Excel file:");

        string? filePath = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Invalid file path or file does not exist.");
            Console.ResetColor();
            return;
        }

        filePath = Path.GetFullPath(filePath);
        Console.WriteLine($"\nLoaded file: {filePath}");
        Console.WriteLine("\nYou can now ask questions about your spreadsheet.");
        Console.WriteLine("Type 'exit' to quit.\n");

        while (true)
        {
            Console.Write("> ");
            string? query = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(query))
                continue;

            if (query.Equals("exit", StringComparison.OrdinalIgnoreCase))
                break;

            try
            {
                Console.WriteLine();
                var result = await plugin.QuerySpreadsheetAsync(filePath, query);

                // Parse and display the result
                var jsonResult = JsonSerializer.Deserialize<JsonElement>(result);

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\n=== Result ===");
                Console.ResetColor();

                if (jsonResult.TryGetProperty("Success", out var success) && success.GetBoolean())
                {
                    if (jsonResult.TryGetProperty("Answer", out var answer))
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Answer: {answer}");
                        Console.ResetColor();
                    }

                    if (jsonResult.TryGetProperty("Formula", out var formula) &&
                        !string.IsNullOrEmpty(formula.GetString()))
                    {
                        Console.WriteLine($"Formula used: {formula}");
                    }

                    if (jsonResult.TryGetProperty("Reasoning", out var reasoning))
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine($"Reasoning: {reasoning}");
                        Console.ResetColor();
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    if (jsonResult.TryGetProperty("Error", out var error))
                    {
                        Console.WriteLine($"Error: {error}");
                    }

                    Console.ResetColor();
                }

                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: {ex.Message}");
                Console.ResetColor();
                logger.LogError(ex, "Error processing query");
            }
        }

        Console.WriteLine("\nGoodbye!");
        
        if (activityPublisher != null)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"📝 Audit log saved to: {activityPublisher.GetLogFilePath()}");
            Console.ResetColor();
        }
    }

    static async Task RunCommandMode(IHost host, string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ssllm <file_path> <query>");
            return;
        }

        var filePath = Path.GetFullPath(args[0]);
        var query = string.Join(" ", args.Skip(1));

        var plugin = host.Services.GetRequiredService<SpreadsheetPlugin>();
        var activityPublisher = host.Services.GetRequiredService<IActivityPublisher>() as FileAndConsoleActivityPublisher;

        try
        {
            var result = await plugin.QuerySpreadsheetAsync(filePath, query);
            
            // Parse and pretty-print the JSON result
            try
            {
                var jsonResult = JsonSerializer.Deserialize<JsonElement>(result);
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                var prettyJson = JsonSerializer.Serialize(jsonResult, options);
                Console.WriteLine(prettyJson);
            }
            catch
            {
                // If JSON parsing fails, print the raw result
                Console.WriteLine(result);
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: {ex.Message}");
            Console.ResetColor();
            
            if (activityPublisher != null)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"📝 Audit log saved to: {activityPublisher.GetLogFilePath()}");
                Console.ResetColor();
            }
            
            Environment.Exit(1);
        }
        
        // Show log file location on success
        if (activityPublisher != null)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"📝 Audit log saved to: {activityPublisher.GetLogFilePath()}");
            Console.ResetColor();
        }
    }
}