using System.Threading.Tasks;

namespace SpreadsheetCLI.Core.Application.Interfaces;

public interface IActivityPublisher
{
    Task PublishAsync(string eventType, object data);
}