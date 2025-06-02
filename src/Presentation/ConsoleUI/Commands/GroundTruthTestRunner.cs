using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using SpreadsheetCLI.Infrastructure.Ai.SemanticKernel.Plugins;
using SpreadsheetCLI.Application.Interfaces;

namespace SpreadsheetCLI.Presentation.ConsoleUI.Commands
{
    public class GroundTruthTestRunner
    {
        private readonly IServiceProvider _serviceProvider;

        public GroundTruthTestRunner(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        public async Task RunTestsAsync(string dataFile, string truthFile, bool useLlmValidation, bool verbose)
        {
            var plugin = _serviceProvider.GetRequiredService<SpreadsheetPlugin>();
            var testValidation = _serviceProvider.GetService<ITestResultValidationService>();

            if (!File.Exists(dataFile))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: Data file not found: {dataFile}");
                Console.ResetColor();
                return;
            }

            if (!File.Exists(truthFile))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: Ground truth file not found: {truthFile}");
                Console.ResetColor();
                return;
            }

            Console.WriteLine($"\n=== Ground Truth Test Runner ===");
            Console.WriteLine($"Data file: {Path.GetFileName(dataFile)}");
            Console.WriteLine($"Truth file: {Path.GetFileName(truthFile)}");
            Console.WriteLine($"Validation: {(useLlmValidation ? "LLM" : "Pattern Matching")}");
            Console.WriteLine();

            // Extract ground truth using the existing ExtractGroundTruth tool
            var groundTruthData = await ExtractGroundTruthAsync(truthFile);
            if (groundTruthData.Count == 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: No ground truth data found");
                Console.ResetColor();
                return;
            }

            Console.WriteLine($"Found {groundTruthData.Count} test cases\n");

            var results = new List<TestResult>();
            var passed = 0;
            var failed = 0;

            for (int i = 0; i < groundTruthData.Count; i++)
            {
                var (question, expectedAnswer) = groundTruthData[i];
                
                Console.Write($"[{i + 1}/{groundTruthData.Count}] Testing: {question.Substring(0, Math.Min(question.Length, 60))}...");
                
                try
                {
                    var queryResult = await plugin.QuerySpreadsheetAsync(dataFile, question);
                    
                    bool isCorrect = false;
                    string? extractedAnswer = null;
                    string? validationReason = null;

                    if (useLlmValidation && testValidation != null)
                    {
                        var validationRequest = new Application.DTOs.TestValidationRequest
                        {
                            Question = question,
                            ExpectedAnswer = expectedAnswer,
                            ActualOutput = queryResult
                        };
                        
                        var validationResult = await testValidation.ValidateTestResultAsync(validationRequest);
                        
                        isCorrect = validationResult.IsCorrect;
                        extractedAnswer = validationResult.ExtractedAnswer;
                        validationReason = validationResult.Explanation;
                    }
                    else
                    {
                        // Simple pattern matching fallback
                        var jsonResult = JsonSerializer.Deserialize<JsonElement>(queryResult);
                        if (jsonResult.TryGetProperty("Answer", out var answer))
                        {
                            extractedAnswer = answer.ToString();
                            isCorrect = NormalizeAnswer(extractedAnswer).Equals(
                                NormalizeAnswer(expectedAnswer), 
                                StringComparison.OrdinalIgnoreCase
                            );
                        }
                    }

                    if (isCorrect)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(" ✓ PASS");
                        Console.ResetColor();
                        passed++;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(" ✗ FAIL");
                        Console.ResetColor();
                        failed++;
                        
                        if (verbose)
                        {
                            Console.WriteLine($"  Expected: {expectedAnswer}");
                            Console.WriteLine($"  Got: {extractedAnswer ?? "(no answer)"}");
                            if (!string.IsNullOrEmpty(validationReason))
                            {
                                Console.WriteLine($"  Reason: {validationReason}");
                            }
                        }
                    }

                    results.Add(new TestResult
                    {
                        Index = i + 1,
                        Question = question,
                        ExpectedAnswer = expectedAnswer,
                        ExtractedAnswer = extractedAnswer,
                        IsCorrect = isCorrect,
                        ValidationReason = validationReason
                    });
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($" ✗ ERROR: {ex.Message}");
                    Console.ResetColor();
                    failed++;
                    
                    results.Add(new TestResult
                    {
                        Index = i + 1,
                        Question = question,
                        ExpectedAnswer = expectedAnswer,
                        IsCorrect = false,
                        Error = ex.Message
                    });
                }
            }

            // Generate report
            var accuracy = (double)passed / groundTruthData.Count * 100;
            
            Console.WriteLine($"\n=== Test Summary ===");
            Console.WriteLine($"Total tests: {groundTruthData.Count}");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Passed: {passed}");
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Failed: {failed}");
            Console.ResetColor();
            Console.WriteLine($"Accuracy: {accuracy:F1}%");

            // Save detailed report
            await SaveTestReportAsync(results, accuracy);
        }

        private async Task<List<(string question, string answer)>> ExtractGroundTruthAsync(string truthFile)
        {
            var results = new List<(string, string)>();
            
            // Run the ExtractGroundTruth executable
            var extractPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scripts", "ExtractGroundTruth.dll");
            
            if (!File.Exists(extractPath))
            {
                // Try to build it
                var projectPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "scripts", "ExtractGroundTruth.csproj");
                if (File.Exists(projectPath))
                {
                    var buildProcess = new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "dotnet",
                            Arguments = $"build \"{projectPath}\" -c Release",
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            UseShellExecute = false
                        }
                    };
                    
                    buildProcess.Start();
                    await buildProcess.WaitForExitAsync();
                }
            }

            var process = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "dotnet",
                    Arguments = $"\"{extractPath}\" \"{truthFile}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false
                }
            };

            process.Start();
            var output = await process.StandardOutput.ReadToEndAsync();
            await process.WaitForExitAsync();

            foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = line.Split("|||");
                if (parts.Length == 2)
                {
                    results.Add((parts[0].Trim(), parts[1].Trim()));
                }
            }

            return results;
        }

        private string NormalizeAnswer(string answer)
        {
            return answer.Replace(",", "")
                        .Replace("$", "")
                        .Replace("%", "")
                        .Trim();
        }

        private async Task SaveTestReportAsync(List<TestResult> results, double accuracy)
        {
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var reportDir = Path.Combine("test-reports", $"test_{timestamp}");
            Directory.CreateDirectory(reportDir);

            var report = new
            {
                Timestamp = DateTime.Now,
                TotalTests = results.Count,
                Passed = results.Count(r => r.IsCorrect),
                Failed = results.Count(r => !r.IsCorrect),
                Accuracy = accuracy,
                Results = results
            };

            var json = JsonSerializer.Serialize(report, new JsonSerializerOptions 
            { 
                WriteIndented = true 
            });
            
            await File.WriteAllTextAsync(Path.Combine(reportDir, "report.json"), json);
            
            Console.WriteLine($"\nDetailed report saved to: {reportDir}/report.json");
        }

        private class TestResult
        {
            public int Index { get; set; }
            public string Question { get; set; } = string.Empty;
            public string ExpectedAnswer { get; set; } = string.Empty;
            public string? ExtractedAnswer { get; set; }
            public bool IsCorrect { get; set; }
            public string? ValidationReason { get; set; }
            public string? Error { get; set; }
        }
    }
}