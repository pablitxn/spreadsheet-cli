# SpreadsheetCLI

A powerful command-line interface tool for querying Excel spreadsheets using natural language powered by OpenAI's GPT-4 and Semantic Kernel. Features an intuitive CLI with multiple commands for interactive queries, batch processing, and ground truth validation.

## Prerequisites

- .NET 9.0 SDK or later
- OpenAI API key
- Linux x64 environment (for pre-built binaries)

##  Installation

### Clone the Repository
```bash
git clone https://github.com/yourusername/spreadsheet-cli.git
cd spreadsheet-cli
```

### Set up OpenAI API Key
```bash
export OPENAI_API_KEY="your-api-key-here"
```

Or add it to `appsettings.json`:
```json
{
  "OpenAI": {
    "ApiKey": "your-api-key-here"
  }
}
```

### Quick Install
```bash
# Run the installation script
./install.sh
```

### Manual Build
```bash
dotnet build -c Release
```

## Usage

### Command Overview

```bash
ssllm [command] [options]
```

Available commands:
- `query` - Query a spreadsheet with natural language
- `interactive` (default) - Start interactive mode
- `browse` - Browse and select Excel files
- `test` - Run ground truth validation tests  
- `batch` - Process multiple queries in batch
- `config` - Manage configuration

### Interactive Mode (Default)
Start the interactive mode by running without arguments:
```bash
ssllm
# or explicitly
ssllm interactive
```

Features:
- File browser for easy file selection
- Query history
- Real-time results with syntax highlighting
- Commands: `exit`, `clear`, `history`, `file`

### Query Command
Run a single query against a spreadsheet:
```bash
ssllm query <file> "<question>"

# With export options
ssllm query data.xlsx "What is the total?" --export markdown --output result.md
ssllm query data.xlsx "What is the average?" --export csv -o result.csv
ssllm query data.xlsx "Show me the max" --export json --verbose
```

### Browse Command  
Browse and select Excel files interactively:
```bash
ssllm browse
ssllm browse --path ./data --filter "*.xlsx"
```

### Test Command
Run ground truth validation tests:
```bash
# Auto-detect test files in standard locations
ssllm test --auto

# Specify files manually
ssllm test --data data.xlsx --truth ground_truth.xlsx

# Options
ssllm test --auto --verbose  # Show detailed output
ssllm test --auto --no-llm   # Use pattern matching instead of LLM
```

### Batch Processing
Process multiple queries efficiently:
```bash
# From text file (one query per line)
ssllm batch data.xlsx --queries queries.txt

# From JSON file
ssllm batch data.xlsx --queries-json queries.json

# With parallel processing
ssllm batch data.xlsx --queries queries.txt --parallel 4

# Custom output directory
ssllm batch data.xlsx --queries queries.txt --output ./results
```

### Configuration Management
Manage application settings:
```bash
# Set configuration values
ssllm config set openai.key "your-api-key"
ssllm config set default.model "gpt-4o"

# Get configuration values  
ssllm config get openai.key

# List all configurations
ssllm config list
```

### Legacy Command Mode
For backwards compatibility:
```bash
ssllm <file_path> "<query>"
```

## Project Structure

```
spreadsheet-cli/
    src/
    Application/        # Business logic and interfaces
       DTOs/          # Data transfer objects
       Interfaces/    # Service contracts
       Services/      # Core services
    Domain/            # Domain entities and business rules
       Entities/      # Domain models
       Enums/         # Domain enumerations
       ValueObjects/  # Value objects
    Infrastructure/    # External concerns (AI, file storage)
       Ai/           # Semantic Kernel integration
       Mocks/        # Mock implementations
       Repositories/ # Data access
    Presentation/     # User interface layer
        ConsoleUI/    # Console presentation
    scripts/              # Helper scripts
    tests/               # Unit tests
    appsettings.json    # Configuration
```

## Architecture

The project follows **Clean Architecture** principles:

- **Domain Layer**: Core business entities and rules
- **Application Layer**: Use cases and business logic
- **Infrastructure Layer**: External integrations (OpenAI, file system)
- **Presentation Layer**: User interface (CLI)

### Key Components

- **SpreadsheetPlugin**: Main entry point for spreadsheet queries
- **SpreadsheetAnalysisService**: AI-powered data analysis
- **AsposeSpreadsheetRepository**: Excel file handling using Aspose.Cells
- **Semantic Kernel**: Orchestrates AI capabilities

## Features

### 🚀 New CLI Interface
- **Modern command structure** with intuitive verbs and options
- **Interactive file browser** for easy Excel file selection
- **Export capabilities** to JSON, CSV, and Markdown formats
- **Batch processing** for running multiple queries efficiently
- **Configuration management** for storing settings
- **Ground truth validation** for testing accuracy

### 🔍 Natural Language Queries
- Ask questions in plain English
- Automatic formula generation
- Intelligent data extraction and analysis
- Support for complex aggregations and calculations

### 📊 Advanced Analytics
- Statistical functions (average, sum, max, min, std dev)
- Filtering and grouping operations
- Percentage calculations
- Ratio and comparison queries

## Example Queries

Here are some example queries you can run on financial spreadsheet data:

```bash
# Basic aggregations
"What is the total Quantity for SecurityID 101121101?"
"What is the average TotalBaseIncome for rows with PaymentType 'DIV'?"
"What is the maximum TotalBaseIncome in the dataset?"

# Complex analysis
"Which SecurityID contributes the highest aggregate TotalBaseIncome?"
"What is the standard deviation of TotalBaseIncome for SecurityID 828806109?"
"What percentage of rows have Quantity greater than 1000?"

# Advanced queries
"What is the ratio of total TotalBaseIncome of ACTUAL rows to DIV rows?"
"Which SecurityGroup derives the largest share of its TotalBaseIncome from DIV payments?"
"For SecurityID 101121101, what is the coefficient of variation of Quantity?"
```

## Testing

### Unit Tests
Run the test suite:
```bash
dotnet test
```

### Ground Truth Validation
Run comprehensive ground truth tests:
```bash
# Using the new CLI
ssllm test --auto

# Using the legacy script
./scripts/test-ground-truth.sh
```

### Other Test Scripts
Available in the `scripts/` directory:
- `test-audit-log.sh`: Test audit logging functionality
- `test-metadata.sh`: Test metadata extraction
- `query.sh`: Quick query testing

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with [Microsoft Semantic Kernel](https://github.com/microsoft/semantic-kernel)
- Excel processing powered by [Aspose.Cells](https://products.aspose.com/cells/net/)
- AI capabilities provided by [OpenAI](https://openai.com/)

## Generated with
This README was Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>