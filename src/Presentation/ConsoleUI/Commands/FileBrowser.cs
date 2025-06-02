using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace SpreadsheetCLI.Presentation.ConsoleUI.Commands
{
    public class FileBrowser
    {
        public async Task<string?> BrowseAsync(string startPath, string filter = "*.xlsx")
        {
            return await Task.Run(() => Browse(startPath, filter));
        }

        private string? Browse(string startPath, string filter)
        {
            string currentPath = Path.GetFullPath(startPath);
            
            while (true)
            {
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("=== File Browser ===");
                Console.ResetColor();
                Console.WriteLine($"Current: {currentPath}");
                Console.WriteLine($"Filter: {filter}\n");

                var entries = new List<BrowserEntry>();
                
                // Add parent directory option
                if (currentPath != "/")
                {
                    entries.Add(new BrowserEntry
                    {
                        Name = "..",
                        FullPath = Path.GetDirectoryName(currentPath) ?? "/",
                        IsDirectory = true
                    });
                }

                // Add directories
                try
                {
                    var dirs = Directory.GetDirectories(currentPath)
                        .OrderBy(d => Path.GetFileName(d))
                        .Select(d => new BrowserEntry
                        {
                            Name = Path.GetFileName(d),
                            FullPath = d,
                            IsDirectory = true
                        });
                    entries.AddRange(dirs);
                }
                catch { }

                // Add files matching filter
                try
                {
                    var files = Directory.GetFiles(currentPath, filter)
                        .OrderBy(f => Path.GetFileName(f))
                        .Select(f => new BrowserEntry
                        {
                            Name = Path.GetFileName(f),
                            FullPath = f,
                            IsDirectory = false,
                            Size = new FileInfo(f).Length
                        });
                    entries.AddRange(files);
                }
                catch { }

                // Display entries
                for (int i = 0; i < entries.Count; i++)
                {
                    var entry = entries[i];
                    
                    if (entry.IsDirectory)
                    {
                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.Write($"{i + 1,3}. [DIR] ");
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{i + 1,3}. [FILE]");
                    }
                    
                    Console.ResetColor();
                    Console.Write($" {entry.Name}");
                    
                    if (!entry.IsDirectory && entry.Size.HasValue)
                    {
                        Console.ForegroundColor = ConsoleColor.DarkGray;
                        Console.Write($" ({FormatFileSize(entry.Size.Value)})");
                        Console.ResetColor();
                    }
                    
                    Console.WriteLine();
                }

                Console.WriteLine("\nEnter number to select, 'q' to quit, or type a path:");
                Console.Write("> ");
                
                var input = Console.ReadLine()?.Trim();
                
                if (string.IsNullOrEmpty(input) || input.Equals("q", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                if (int.TryParse(input, out int selection) && selection > 0 && selection <= entries.Count)
                {
                    var selected = entries[selection - 1];
                    
                    if (selected.IsDirectory)
                    {
                        currentPath = selected.FullPath;
                    }
                    else
                    {
                        return selected.FullPath;
                    }
                }
                else if (Directory.Exists(input))
                {
                    currentPath = Path.GetFullPath(input);
                }
                else if (File.Exists(input))
                {
                    return Path.GetFullPath(input);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Invalid selection. Press any key to continue...");
                    Console.ResetColor();
                    Console.ReadKey(true);
                }
            }
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }

            return $"{len:0.##} {sizes[order]}";
        }

        private class BrowserEntry
        {
            public string Name { get; set; } = string.Empty;
            public string FullPath { get; set; } = string.Empty;
            public bool IsDirectory { get; set; }
            public long? Size { get; set; }
        }
    }
}