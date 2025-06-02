using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Plugins;

namespace SpreadsheetCLI.Presentation.ConsoleUI.Commands
{
    public class BatchQueryProcessor
    {
        private readonly IServiceProvider _serviceProvider;

        public BatchQueryProcessor(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        public async Task ProcessBatchAsync(string filePath, List<string> queries, string outputDir, int parallelism)
        {
            var plugin = _serviceProvider.GetRequiredService<SpreadsheetPlugin>();
            
            // Ensure output directory exists
            Directory.CreateDirectory(outputDir);
            
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var batchDir = Path.Combine(outputDir, $"batch_{timestamp}");
            Directory.CreateDirectory(batchDir);

            Console.WriteLine($"\nProcessing {queries.Count} queries...");
            Console.WriteLine($"Output directory: {batchDir}\n");

            var results = new List<BatchResult>();
            var semaphore = new System.Threading.SemaphoreSlim(parallelism);
            var tasks = new List<Task>();

            for (int i = 0; i < queries.Count; i++)
            {
                var index = i;
                var query = queries[i];
                
                if (string.IsNullOrWhiteSpace(query))
                    continue;

                var task = Task.Run(async () =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        var startTime = DateTime.Now;
                        Console.WriteLine($"[{index + 1}/{queries.Count}] Processing: {query.Substring(0, Math.Min(query.Length, 50))}...");

                        try
                        {
                            var result = await plugin.QuerySpreadsheetAsync(filePath, query);
                            var elapsed = DateTime.Now - startTime;
                            
                            var batchResult = new BatchResult
                            {
                                Index = index + 1,
                                Query = query,
                                Success = true,
                                Result = result,
                                ProcessingTime = elapsed.TotalSeconds
                            };
                            
                            lock (results)
                            {
                                results.Add(batchResult);
                            }

                            // Save individual result
                            var resultFile = Path.Combine(batchDir, $"query_{index + 1:D4}.json");
                            await File.WriteAllTextAsync(resultFile, result);

                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine($"[{index + 1}/{queries.Count}] ✓ Completed in {elapsed.TotalSeconds:F1}s");
                            Console.ResetColor();
                        }
                        catch (Exception ex)
                        {
                            var elapsed = DateTime.Now - startTime;
                            
                            var batchResult = new BatchResult
                            {
                                Index = index + 1,
                                Query = query,
                                Success = false,
                                Error = ex.Message,
                                ProcessingTime = elapsed.TotalSeconds
                            };
                            
                            lock (results)
                            {
                                results.Add(batchResult);
                            }

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"[{index + 1}/{queries.Count}] ✗ Failed: {ex.Message}");
                            Console.ResetColor();
                        }
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });

                tasks.Add(task);
            }

            await Task.WhenAll(tasks);

            // Generate summary report
            await GenerateSummaryReport(batchDir, results);
            
            Console.WriteLine($"\n=== Batch Processing Complete ===");
            Console.WriteLine($"Total queries: {results.Count}");
            Console.WriteLine($"Successful: {results.Count(r => r.Success)}");
            Console.WriteLine($"Failed: {results.Count(r => !r.Success)}");
            Console.WriteLine($"Total time: {results.Sum(r => r.ProcessingTime):F1}s");
            Console.WriteLine($"Results saved to: {batchDir}");
        }

        private async Task GenerateSummaryReport(string batchDir, List<BatchResult> results)
        {
            var summary = new
            {
                Timestamp = DateTime.Now,
                TotalQueries = results.Count,
                Successful = results.Count(r => r.Success),
                Failed = results.Count(r => !r.Success),
                TotalProcessingTime = results.Sum(r => r.ProcessingTime),
                AverageProcessingTime = results.Average(r => r.ProcessingTime),
                Results = results.OrderBy(r => r.Index).Select(r => new
                {
                    r.Index,
                    r.Query,
                    r.Success,
                    r.ProcessingTime,
                    r.Error,
                    Answer = r.Success ? ExtractAnswer(r.Result) : null
                })
            };

            var summaryJson = JsonSerializer.Serialize(summary, new JsonSerializerOptions 
            { 
                WriteIndented = true 
            });
            
            await File.WriteAllTextAsync(Path.Combine(batchDir, "summary.json"), summaryJson);

            // Also create a CSV report
            var csvLines = new List<string>
            {
                "Index,Query,Success,ProcessingTime,Answer,Error"
            };

            foreach (var result in results.OrderBy(r => r.Index))
            {
                var answer = result.Success ? ExtractAnswer(result.Result) : "";
                var error = result.Success ? "" : result.Error;
                
                csvLines.Add($"{result.Index},\"{result.Query}\",{result.Success},{result.ProcessingTime:F2},\"{answer}\",\"{error}\"");
            }

            await File.WriteAllLinesAsync(Path.Combine(batchDir, "summary.csv"), csvLines);
        }

        private string? ExtractAnswer(string? jsonResult)
        {
            if (string.IsNullOrEmpty(jsonResult))
                return null;

            try
            {
                var json = JsonSerializer.Deserialize<JsonElement>(jsonResult);
                if (json.TryGetProperty("Answer", out var answer))
                {
                    return answer.ToString();
                }
            }
            catch { }

            return null;
        }

        private class BatchResult
        {
            public int Index { get; set; }
            public string Query { get; set; } = string.Empty;
            public bool Success { get; set; }
            public string? Result { get; set; }
            public string? Error { get; set; }
            public double ProcessingTime { get; set; }
        }
    }
}