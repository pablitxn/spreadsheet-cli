using System;
using System.IO;
using System.Threading.Tasks;
using SpreadsheetCLI.Core.Application.Interfaces;

namespace SpreadsheetCLI.Infrastructure.Mocks;

public class LocalFileStorageService : IFileStorageService
{
    public Task<Stream> GetFileAsync(string fileName)
    {
        // For CLI, we'll just read from the file system directly
        if (!File.Exists(fileName))
        {
            throw new FileNotFoundException($"File not found: {fileName}");
        }
        
        return Task.FromResult<Stream>(File.OpenRead(fileName));
    }

}