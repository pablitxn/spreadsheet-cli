using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading.Tasks;

namespace SpreadsheetCLI.Infrastructure.Services
{
    public interface IFileHashService
    {
        Task<string> CalculateHashAsync(string filePath);
        bool ValidateHash(string filePath, string expectedHash);
    }

    public class FileHashService : IFileHashService
    {
        public async Task<string> CalculateHashAsync(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found: {filePath}");
            }

            using var sha256 = SHA256.Create();
            using var stream = File.OpenRead(filePath);
            
            var hashBytes = await sha256.ComputeHashAsync(stream);
            return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
        }

        public bool ValidateHash(string filePath, string expectedHash)
        {
            try
            {
                var actualHash = CalculateHashAsync(filePath).GetAwaiter().GetResult();
                return string.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
    }
}