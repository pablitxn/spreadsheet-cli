using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using SpreadsheetCLI.Application.Interfaces;

namespace SpreadsheetCLI.Infrastructure.Mocks;

public class FileAndConsoleActivityPublisher : IActivityPublisher
{
    private readonly string _logFilePath;
    private readonly object _fileLock = new object();

    public FileAndConsoleActivityPublisher()
    {
        // Create a unique log file name with timestamp
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
        _logFilePath = Path.Combine(Directory.GetCurrentDirectory(), $"audit_log_{timestamp}.txt");
        
        // Write header to file
        WriteToFile($"=== Spreadsheet CLI Audit Log ===");
        WriteToFile($"Started at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        WriteToFile($"Log file: {_logFilePath}");
        WriteToFile(new string('=', 80));
        WriteToFile("");
    }

    public Task PublishAsync(string eventType, object data)
    {
        var timestamp = DateTime.UtcNow.ToString("HH:mm:ss.fff");
        var jsonData = JsonSerializer.Serialize(data, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        // Format the output
        var formattedOutput = $"[{timestamp}] {eventType}:\n{jsonData}\n";

        // Write to console with colors
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.Write($"[{timestamp}] ");
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine($"{eventType}:");
        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine(jsonData);
        Console.ResetColor();
        Console.WriteLine();

        // Write to file
        WriteToFile(formattedOutput);

        return Task.CompletedTask;
    }

    private void WriteToFile(string content)
    {
        lock (_fileLock)
        {
            try
            {
                File.AppendAllText(_logFilePath, content + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error writing to log file: {ex.Message}");
                Console.ResetColor();
            }
        }
    }

    public string GetLogFilePath() => _logFilePath;
}