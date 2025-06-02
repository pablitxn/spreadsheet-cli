using System;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace SpreadsheetCLI.Presentation.ConsoleUI.Commands
{
    public class ResultExporter
    {
        public async Task<string> ExportAsync(JsonElement result, string format)
        {
            return format.ToLower() switch
            {
                "json" => await ExportJsonAsync(result),
                "csv" => await ExportCsvAsync(result),
                "markdown" or "md" => await ExportMarkdownAsync(result),
                _ => throw new ArgumentException($"Unsupported export format: {format}")
            };
        }

        private Task<string> ExportJsonAsync(JsonElement result)
        {
            var options = new JsonSerializerOptions { WriteIndented = true };
            return Task.FromResult(JsonSerializer.Serialize(result, options));
        }

        private Task<string> ExportCsvAsync(JsonElement result)
        {
            var csv = new StringBuilder();
            
            // Header
            csv.AppendLine("Field,Value");
            
            // Add fields
            if (result.TryGetProperty("Success", out var success))
                csv.AppendLine($"Success,{success}");
            
            if (result.TryGetProperty("Answer", out var answer))
                csv.AppendLine($"Answer,\"{answer}\"");
            
            if (result.TryGetProperty("Formula", out var formula))
                csv.AppendLine($"Formula,\"{formula}\"");
            
            if (result.TryGetProperty("Reasoning", out var reasoning))
                csv.AppendLine($"Reasoning,\"{reasoning}\"");
            
            if (result.TryGetProperty("Error", out var error))
                csv.AppendLine($"Error,\"{error}\"");
            
            if (result.TryGetProperty("ProcessingTime", out var time))
                csv.AppendLine($"ProcessingTime,{time}");

            return Task.FromResult(csv.ToString());
        }

        private Task<string> ExportMarkdownAsync(JsonElement result)
        {
            var md = new StringBuilder();
            
            md.AppendLine("# Query Result\n");
            
            if (result.TryGetProperty("Success", out var success))
            {
                md.AppendLine($"**Status**: {(success.GetBoolean() ? "✓ Success" : "✗ Failed")}\n");
            }
            
            if (result.TryGetProperty("Answer", out var answer))
            {
                md.AppendLine("## Answer");
                md.AppendLine($"{answer}\n");
            }
            
            if (result.TryGetProperty("Formula", out var formula) && !string.IsNullOrEmpty(formula.GetString()))
            {
                md.AppendLine("## Formula");
                md.AppendLine("```");
                md.AppendLine(formula.GetString());
                md.AppendLine("```\n");
            }
            
            if (result.TryGetProperty("Reasoning", out var reasoning))
            {
                md.AppendLine("## Reasoning");
                md.AppendLine($"{reasoning}\n");
            }
            
            if (result.TryGetProperty("ExecutionPlan", out var plan))
            {
                md.AppendLine("## Execution Plan");
                md.AppendLine("```json");
                md.AppendLine(JsonSerializer.Serialize(plan, new JsonSerializerOptions { WriteIndented = true }));
                md.AppendLine("```\n");
            }
            
            if (result.TryGetProperty("Error", out var error))
            {
                md.AppendLine("## Error");
                md.AppendLine($"{error}\n");
            }
            
            if (result.TryGetProperty("ProcessingTime", out var time))
            {
                md.AppendLine($"*Processing time: {time}s*");
            }

            return Task.FromResult(md.ToString());
        }
    }
}