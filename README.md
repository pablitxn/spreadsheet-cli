# SpreadsheetCLI

A powerful command-line interface tool for querying Excel spreadsheets using natural language powered by OpenAI's `o4-mini` and Semantic Kernel.

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

### Build the Project
```bash
dotnet build
```

## Usage

### Interactive Mode
Simply run the executable without arguments:
```bash
./bin/Debug/net9.0/linux-x64/ssllm
```

Then follow the prompts to:
1. Enter your Excel file path
2. Ask questions about your data
3. Type 'exit' to quit

### Command Mode
For scripting and automation:
```bash
./bin/Debug/net9.0/linux-x64/ssllm <file_path> "<query>"
```

Example:
```bash
./bin/Debug/net9.0/linux-x64/ssllm expanded_dataset_moved.xlsx "What is the total Quantity for SecurityID 101121101?"
```

### Using Helper Scripts

#### Quick Query (with build)
```bash
./scripts/query.sh "What is the average TotalBaseIncome for DIV payments?"
```

#### Fast Query (no build)
```bash
./scripts/query-fast.sh "Which SecurityID has the highest aggregate TotalBaseIncome?"
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

Run the test suite:
```bash
dotnet test
```

Test scripts are available in the `scripts/` directory:
- `test-audit-log.sh`: Test audit logging functionality
- `test-metadata.sh`: Test metadata extraction

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