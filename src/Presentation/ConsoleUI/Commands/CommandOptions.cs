using CommandLine;
using System.Collections.Generic;

namespace SpreadsheetCLI.Presentation.ConsoleUI.Commands
{
    [Verb("query", HelpText = "Query a spreadsheet with natural language")]
    public class QueryOptions
    {
        [Value(0, MetaName = "file", Required = true, HelpText = "Path to the Excel file")]
        public string FilePath { get; set; } = string.Empty;

        [Value(1, MetaName = "query", Required = true, HelpText = "Natural language query")]
        public string Query { get; set; } = string.Empty;

        [Option('e', "export", Required = false, HelpText = "Export format (json, csv, markdown)")]
        public string? ExportFormat { get; set; }

        [Option('o', "output", Required = false, HelpText = "Output file path for export")]
        public string? OutputPath { get; set; }

        [Option('v', "verbose", Required = false, HelpText = "Enable verbose output")]
        public bool Verbose { get; set; }
    }

    [Verb("interactive", HelpText = "Start interactive mode")]
    public class InteractiveOptions
    {
        [Option('f', "file", Required = false, HelpText = "Preload an Excel file")]
        public string? FilePath { get; set; }

        [Option('h', "history", Required = false, HelpText = "Load query history from file")]
        public string? HistoryFile { get; set; }
    }

    [Verb("browse", HelpText = "Browse and select Excel files")]
    public class BrowseOptions
    {
        [Option('p', "path", Required = false, HelpText = "Starting directory path")]
        public string Path { get; set; } = ".";

        [Option('f', "filter", Required = false, HelpText = "File filter pattern")]
        public string Filter { get; set; } = "*.xlsx";
    }

    [Verb("test", HelpText = "Run ground truth validation tests")]
    public class TestOptions
    {
        [Option('d', "data", Required = false, HelpText = "Data file path")]
        public string? DataFile { get; set; }

        [Option('t', "truth", Required = false, HelpText = "Ground truth file path")]
        public string? TruthFile { get; set; }

        [Option('a', "auto", Required = false, HelpText = "Auto-detect files in standard locations")]
        public bool Auto { get; set; }

        [Option('l', "llm", Required = false, HelpText = "Use LLM validation")]
        public bool UseLlm { get; set; } = true;

        [Option('v', "verbose", Required = false, HelpText = "Verbose output")]
        public bool Verbose { get; set; }
    }

    [Verb("batch", HelpText = "Process multiple queries in batch")]
    public class BatchOptions
    {
        [Value(0, MetaName = "file", Required = true, HelpText = "Path to the Excel file")]
        public string FilePath { get; set; } = string.Empty;

        [Option('q', "queries", Required = false, HelpText = "Text file with queries (one per line)")]
        public string? QueriesFile { get; set; }

        [Option('j', "queries-json", Required = false, HelpText = "JSON file with queries")]
        public string? QueriesJsonFile { get; set; }

        [Option('o', "output", Required = false, HelpText = "Output directory for results")]
        public string OutputDir { get; set; } = "./batch-results";

        [Option('p', "parallel", Required = false, HelpText = "Number of parallel queries")]
        public int Parallel { get; set; } = 1;
    }

    [Verb("config", HelpText = "Manage configuration")]
    public class ConfigOptions
    {
        [Value(0, MetaName = "action", Required = true, HelpText = "Action: get, set, list")]
        public string Action { get; set; } = string.Empty;

        [Value(1, MetaName = "key", Required = false, HelpText = "Configuration key")]
        public string? Key { get; set; }

        [Value(2, MetaName = "value", Required = false, HelpText = "Configuration value")]
        public string? Value { get; set; }
    }
}