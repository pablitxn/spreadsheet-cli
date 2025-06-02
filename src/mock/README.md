# Simplified Spreadsheet Plugin Architecture

This directory contains a simplified version of the spreadsheet plugin with only 3 files:

## Files

1. **ISpreadsheetService.cs** - The unified interface that combines all spreadsheet operations
2. **SpreadsheetService.cs** - The implementation with all business logic consolidated
3. **SpreadsheetPlugin.cs** - The Semantic Kernel plugin that uses the service

## Key Changes from Original Architecture

### Before (Hexagonal Architecture)
- Multiple layers: Domain, Application, Infrastructure
- Separate interfaces for each concern
- Repository pattern
- Multiple services and DTOs spread across directories

### After (Simplified)
- Single interface for all spreadsheet operations
- One implementation file with all logic
- DTOs included in the interface file
- Direct usage without repository abstraction

## Usage Example

```csharp
// Register services
services.AddSingleton<ISpreadsheetService, SpreadsheetService>();
services.AddSingleton<SpreadsheetPlugin>();

// In your Semantic Kernel setup
var kernel = Kernel.CreateBuilder()
    .AddOpenAIChatCompletion("model-name", "api-key")
    .Build();

var plugin = serviceProvider.GetRequiredService<SpreadsheetPlugin>();
kernel.ImportPluginFromObject(plugin, "spreadsheet");

// Use the plugin
var result = await kernel.InvokeAsync(
    "spreadsheet",
    "query_spreadsheet",
    new()
    {
        ["filePath"] = "path/to/file.xlsx",
        ["query"] = "What is the total sales?",
        ["sheetName"] = "Sheet1"
    }
);
```

## Dependencies Required

- Aspose.Cells
- Microsoft.SemanticKernel
- Microsoft.Extensions.Logging

## Benefits of This Approach

1. **Simplicity** - Only 3 files to maintain
2. **Portability** - Easy to copy to other projects
3. **Clear Dependencies** - All dependencies visible in one place
4. **No Architecture Overhead** - Direct implementation without layers

## Migration Guide

To use this in another application:

1. Copy these 3 files to your project
2. Add the required NuGet packages
3. Register ISpreadsheetService and SpreadsheetPlugin in DI
4. Configure Semantic Kernel with the plugin
5. Start querying spreadsheets!

## Note

This simplified version maintains all the core functionality:
- Natural language query processing
- Dynamic spreadsheet analysis
- Formula execution
- Multiple format detection
- Comprehensive metadata extraction

The main difference is the architectural approach - everything is consolidated for maximum simplicity and portability.