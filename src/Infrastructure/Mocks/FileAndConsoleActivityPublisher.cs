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
        // Check if custom audit log directory is set via environment variable
        var auditLogDir = Environment.GetEnvironmentVariable("AUDIT_LOG_DIR");
        if (string.IsNullOrEmpty(auditLogDir))
        {
            auditLogDir = Path.Combine(Directory.GetCurrentDirectory(), "logs", "audit");
        }
        
        // Create directory if it doesn't exist
        Directory.CreateDirectory(auditLogDir);
        
        // Create a unique log file name with timestamp
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
        var logFileName = $"audit_{timestamp}.json";
        _logFilePath = Path.Combine(auditLogDir, logFileName);
    }

    public Task PublishAsync(string eventType, object data)
    {
        var timestamp = DateTime.UtcNow;
        
        // Create a structured log entry
        var logEntry = new
        {
            Timestamp = timestamp.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"),
            EventType = eventType,
            Data = data
        };

        var jsonData = JsonSerializer.Serialize(logEntry, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        // Write to console with colors
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.Write($"[{timestamp:HH:mm:ss.fff}] ");
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine($"{eventType}:");
        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine(JsonSerializer.Serialize(data, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        }));
        Console.ResetColor();
        Console.WriteLine();

        // Write to file as JSON
        WriteToFile(jsonData);

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