using System.IO;
using System.Threading.Tasks;

namespace SpreadsheetCLI.Core.Application.Interfaces;

public interface IFileStorageService
{
    Task<Stream> GetFileAsync(string fileName);
}