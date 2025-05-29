namespace SpreadsheetCLI.Presentation.ConsoleUI;

using System;
using System.Threading.Tasks;
using SpreadsheetCLI.Application.Interfaces;

public class ConsoleActivityPublisher : IActivityPublisher
{
    public Task PublishAsync(string eventType, object data)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss");
        var color = Console.ForegroundColor;
        
        try
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write($"[{timestamp}] ");
            
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"{eventType}: {data}");
        }
        finally
        {
            Console.ForegroundColor = color;
        }
        
        return Task.CompletedTask;
    }
    
    public string GetLogFilePath()
    {
        // Console publisher doesn't write to file
        return string.Empty;
    }
}