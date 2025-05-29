using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace SpreadsheetCLI.Infrastructure.Services;

public class FileLoggerService
{
    private readonly string _logDirectory;
    private readonly string _sessionId;
    private readonly object _lockObject = new();

    public FileLoggerService()
    {
        _logDirectory = Path.Combine(Directory.GetCurrentDirectory(), "logs");
        _sessionId = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
        
        Directory.CreateDirectory(_logDirectory);
    }

    public async Task LogDebugAsync(string eventName, object data)
    {
        var logEntry = new
        {
            Timestamp = DateTime.UtcNow,
            SessionId = _sessionId,
            EventName = eventName,
            Data = data
        };

        var json = JsonSerializer.Serialize(logEntry, new JsonSerializerOptions 
        { 
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        var fileName = $"debug_{_sessionId}.log";
        var filePath = Path.Combine(_logDirectory, fileName);

        lock (_lockObject)
        {
            File.AppendAllText(filePath, json + Environment.NewLine + new string('-', 80) + Environment.NewLine);
        }

        await Task.CompletedTask;
    }
}