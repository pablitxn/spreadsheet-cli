using System.IO;
using System.Threading.Tasks;

namespace SpreadsheetCLI.Application.Interfaces;

public interface IFileStorageService
{
    Task<Stream> GetFileAsync(string fileName);
}