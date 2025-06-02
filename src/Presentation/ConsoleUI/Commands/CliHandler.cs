using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Plugins;
using SpreadsheetCLI.Application.Interfaces;

namespace SpreadsheetCLI.Presentation.ConsoleUI.Commands
{
    public class CliHandler
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly ILogger<CliHandler> _logger;

        public CliHandler(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
            _logger = serviceProvider.GetRequiredService<ILogger<CliHandler>>();
        }

        public async Task<int> HandleQueryAsync(QueryOptions options)
        {
            try
            {
                var plugin = _serviceProvider.GetRequiredService<SpreadsheetPlugin>();
                var filePath = Path.GetFullPath(options.FilePath);

                if (!File.Exists(filePath))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error: File not found: {filePath}");
                    Console.ResetColor();
                    return 1;
                }

                if (options.Verbose)
                {
                    Console.WriteLine($"Processing query on file: {filePath}");
                    Console.WriteLine($"Query: {options.Query}");
                }

                var result = await plugin.QuerySpreadsheetAsync(filePath, options.Query);
                var jsonResult = JsonSerializer.Deserialize<JsonElement>(result);

                if (options.ExportFormat != null)
                {
                    await ExportResultAsync(jsonResult, options.ExportFormat, options.OutputPath);
                }
                else
                {
                    DisplayResult(jsonResult, options.Verbose);
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: {ex.Message}");
                Console.ResetColor();
                _logger.LogError(ex, "Error in query command");
                return 1;
            }
        }

        public async Task<int> HandleInteractiveAsync(InteractiveOptions options)
        {
            var plugin = _serviceProvider.GetRequiredService<SpreadsheetPlugin>();
            var activityPublisher = _serviceProvider.GetRequiredService<IActivityPublisher>();
            
            Console.Clear();
            PrintBanner();

            string? filePath = options.FilePath;
            
            if (string.IsNullOrEmpty(filePath))
            {
                filePath = await SelectFileInteractive();
                if (string.IsNullOrEmpty(filePath))
                    return 0;
            }

            filePath = Path.GetFullPath(filePath);
            
            if (!File.Exists(filePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: File not found: {filePath}");
                Console.ResetColor();
                return 1;
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\n✓ Loaded: {Path.GetFileName(filePath)}");
            Console.ResetColor();
            
            var history = new List<string>();
            if (!string.IsNullOrEmpty(options.HistoryFile) && File.Exists(options.HistoryFile))
            {
                history.AddRange(File.ReadAllLines(options.HistoryFile));
            }

            Console.WriteLine("\nType your queries below. Commands:");
            Console.WriteLine("  'exit' or 'quit' - Exit the program");
            Console.WriteLine("  'clear' - Clear the screen");
            Console.WriteLine("  'history' - Show query history");
            Console.WriteLine("  'file' - Change Excel file\n");

            
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write("query> ");
                Console.ResetColor();
                
                string? input = Console.ReadLine();
                
                if (string.IsNullOrWhiteSpace(input))
                    continue;

                if (input.Equals("exit", StringComparison.OrdinalIgnoreCase) || 
                    input.Equals("quit", StringComparison.OrdinalIgnoreCase))
                    break;

                if (input.Equals("clear", StringComparison.OrdinalIgnoreCase))
                {
                    Console.Clear();
                    PrintBanner();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"✓ Loaded: {Path.GetFileName(filePath)}");
                    Console.ResetColor();
                    continue;
                }

                if (input.Equals("history", StringComparison.OrdinalIgnoreCase))
                {
                    ShowHistory(history);
                    continue;
                }

                if (input.Equals("file", StringComparison.OrdinalIgnoreCase))
                {
                    var newFile = await SelectFileInteractive();
                    if (!string.IsNullOrEmpty(newFile))
                    {
                        filePath = Path.GetFullPath(newFile);
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"\n✓ Loaded: {Path.GetFileName(filePath)}");
                        Console.ResetColor();
                    }
                    continue;
                }

                history.Add(input);
                
                try
                {
                    var result = await plugin.QuerySpreadsheetAsync(filePath, input);
                    var jsonResult = JsonSerializer.Deserialize<JsonElement>(result);
                    DisplayResult(jsonResult, verbose: false);
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error: {ex.Message}");
                    Console.ResetColor();
                }
            }

            if (!string.IsNullOrEmpty(options.HistoryFile))
            {
                File.WriteAllLines(options.HistoryFile, history);
            }

            Console.WriteLine("\nGoodbye! 👋");
            return 0;
        }

        public async Task<int> HandleBrowseAsync(BrowseOptions options)
        {
            var browser = new FileBrowser();
            var selectedFile = await browser.BrowseAsync(options.Path, options.Filter);
            
            if (!string.IsNullOrEmpty(selectedFile))
            {
                Console.WriteLine(selectedFile);
                return 0;
            }
            
            return 1;
        }

        public async Task<int> HandleTestAsync(TestOptions options)
        {
            try
            {
                string? dataFile = options.DataFile;
                string? truthFile = options.TruthFile;

                if (options.Auto)
                {
                    dataFile = "./scripts/dataset/expanded_dataset_moved.xlsx";
                    truthFile = "./scripts/dataset/ground_truth_expanded_dataset_moved.xlsx";
                }

                if (string.IsNullOrEmpty(dataFile) || string.IsNullOrEmpty(truthFile))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: Data file and truth file are required");
                    Console.ResetColor();
                    return 1;
                }

                var testRunner = new GroundTruthTestRunner(_serviceProvider);
                await testRunner.RunTestsAsync(dataFile, truthFile, options.UseLlm, options.Verbose);
                
                return 0;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: {ex.Message}");
                Console.ResetColor();
                return 1;
            }
        }

        public async Task<int> HandleBatchAsync(BatchOptions options)
        {
            try
            {
                var queries = new List<string>();

                if (!string.IsNullOrEmpty(options.QueriesFile))
                {
                    queries.AddRange(File.ReadAllLines(options.QueriesFile));
                }
                else if (!string.IsNullOrEmpty(options.QueriesJsonFile))
                {
                    var json = File.ReadAllText(options.QueriesJsonFile);
                    var queryArray = JsonSerializer.Deserialize<string[]>(json);
                    if (queryArray != null)
                        queries.AddRange(queryArray);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: Queries file is required (--queries or --queries-json)");
                    Console.ResetColor();
                    return 1;
                }

                var batchProcessor = new BatchQueryProcessor(_serviceProvider);
                await batchProcessor.ProcessBatchAsync(
                    options.FilePath, 
                    queries, 
                    options.OutputDir, 
                    options.Parallel
                );

                return 0;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: {ex.Message}");
                Console.ResetColor();
                return 1;
            }
        }

        public async Task<int> HandleConfigAsync(ConfigOptions options)
        {
            var configManager = new ConfigurationManager();

            switch (options.Action.ToLower())
            {
                case "get":
                    if (string.IsNullOrEmpty(options.Key))
                    {
                        Console.WriteLine("Error: Key is required for 'get' action");
                        return 1;
                    }
                    var value = await configManager.GetAsync(options.Key);
                    Console.WriteLine(value ?? "(not set)");
                    break;

                case "set":
                    if (string.IsNullOrEmpty(options.Key) || string.IsNullOrEmpty(options.Value))
                    {
                        Console.WriteLine("Error: Key and value are required for 'set' action");
                        return 1;
                    }
                    await configManager.SetAsync(options.Key, options.Value);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"✓ Configuration updated: {options.Key}");
                    Console.ResetColor();
                    break;

                case "list":
                    var configs = await configManager.ListAsync();
                    foreach (var config in configs)
                    {
                        Console.WriteLine($"{config.Key}: {config.Value}");
                    }
                    break;

                default:
                    Console.WriteLine($"Error: Unknown action '{options.Action}'");
                    return 1;
            }

            return 0;
        }

        private void DisplayResult(JsonElement result, bool verbose)
        {
            if (result.TryGetProperty("Success", out var success) && success.GetBoolean())
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\n✓ Success");
                Console.ResetColor();

                if (result.TryGetProperty("Answer", out var answer))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"\nAnswer: {answer}");
                    Console.ResetColor();
                }

                if (verbose)
                {
                    if (result.TryGetProperty("Formula", out var formula) && 
                        !string.IsNullOrEmpty(formula.GetString()))
                    {
                        Console.WriteLine($"\nFormula: {formula}");
                    }

                    if (result.TryGetProperty("Reasoning", out var reasoning))
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine($"\nReasoning: {reasoning}");
                        Console.ResetColor();
                    }

                    if (result.TryGetProperty("ExecutionPlan", out var plan))
                    {
                        Console.WriteLine("\nExecution Plan:");
                        Console.WriteLine(JsonSerializer.Serialize(plan, new JsonSerializerOptions { WriteIndented = true }));
                    }
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\n✗ Failed");
                
                if (result.TryGetProperty("Error", out var error))
                {
                    Console.WriteLine($"\nError: {error}");
                }
                Console.ResetColor();
            }
        }

        private async Task ExportResultAsync(JsonElement result, string format, string? outputPath)
        {
            var exporter = new ResultExporter();
            var exported = await exporter.ExportAsync(result, format);
            
            if (string.IsNullOrEmpty(outputPath))
            {
                Console.WriteLine(exported);
            }
            else
            {
                await File.WriteAllTextAsync(outputPath, exported);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ Result exported to: {outputPath}");
                Console.ResetColor();
            }
        }

        private async Task<string?> SelectFileInteractive()
        {
            var browser = new FileBrowser();
            Console.WriteLine("\nSelect an Excel file:");
            return await browser.BrowseAsync(".", "*.xlsx");
        }

        private void ShowHistory(List<string> history)
        {
            if (history.Count == 0)
            {
                Console.WriteLine("No query history available.");
                return;
            }

            Console.WriteLine("\n=== Query History ===");
            for (int i = 0; i < history.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {history[i]}");
            }
            Console.WriteLine();
        }

        private void PrintBanner()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(@"
╔═══════════════════════════════════════════╗
║          SpreadsheetCLI v2.0              ║
║   Natural Language Excel Query Tool       ║
╚═══════════════════════════════════════════╝");
            Console.ResetColor();
        }
    }
}