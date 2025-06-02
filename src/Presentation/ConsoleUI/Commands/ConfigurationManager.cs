using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace SpreadsheetCLI.Presentation.ConsoleUI.Commands
{
    public class ConfigurationManager
    {
        private readonly string _configPath;
        private Dictionary<string, string> _config;

        public ConfigurationManager()
        {
            var homeDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var configDir = Path.Combine(homeDir, ".spreadsheet-cli");
            Directory.CreateDirectory(configDir);
            
            _configPath = Path.Combine(configDir, "config.json");
            _config = LoadConfig();
        }

        public Task<string?> GetAsync(string key)
        {
            _config.TryGetValue(key, out var value);
            return Task.FromResult(value);
        }

        public async Task SetAsync(string key, string value)
        {
            _config[key] = value;
            await SaveConfigAsync();
        }

        public Task<IEnumerable<KeyValuePair<string, string>>> ListAsync()
        {
            return Task.FromResult<IEnumerable<KeyValuePair<string, string>>>(_config);
        }

        private Dictionary<string, string> LoadConfig()
        {
            if (File.Exists(_configPath))
            {
                try
                {
                    var json = File.ReadAllText(_configPath);
                    return JsonSerializer.Deserialize<Dictionary<string, string>>(json) 
                           ?? new Dictionary<string, string>();
                }
                catch
                {
                    return new Dictionary<string, string>();
                }
            }

            return new Dictionary<string, string>();
        }

        private async Task SaveConfigAsync()
        {
            var json = JsonSerializer.Serialize(_config, new JsonSerializerOptions 
            { 
                WriteIndented = true 
            });
            await File.WriteAllTextAsync(_configPath, json);
        }
    }
}