using System.Threading.Tasks;

namespace SpreadsheetCLI.Application.Interfaces;

public interface IActivityPublisher
{
    Task PublishAsync(string eventType, object data);
    string GetLogFilePath();
}