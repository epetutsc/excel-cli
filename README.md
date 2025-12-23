# Excel CLI

A powerful command-line tool for reading and modifying Excel files using ClosedXML. This tool processes local Excel files (.xlsx format) and provides various commands for data manipulation.

## Features

- 📖 **Read Operations**: Read file info, list sheets, read cells and ranges
- ✏️ **Write Operations**: Write to cells, insert formulas
- 📋 **Sheet Management**: Create, delete, rename, and copy worksheets
- 🔍 **Search**: Find values in worksheets
- 📤 **Export**: Export sheets to CSV or JSON
- 📥 **Import**: Import data from CSV or JSON files

## Prerequisites

- .NET 10.0 or later

## Installation

```bash
git clone https://github.com/epetutsc/excel-cli.git
cd excel-cli
dotnet build
```

## Usage

Run commands with:

```bash
dotnet run --project src/ExcelCli/ExcelCli.csproj -- <command> [options]
```

### Available Commands

#### Read Operations

**read-file** - Display file information and sheet summary
```bash
dotnet run -- read-file --path data.xlsx
```

**list-sheets** - List all worksheets
```bash
dotnet run -- list-sheets --path data.xlsx
```

**read-cell** - Read a specific cell value
```bash
dotnet run -- read-cell --path data.xlsx --sheet "Sheet1" --cell A1
```

**read-range** - Read a range of cells
```bash
dotnet run -- read-range --path data.xlsx --sheet "Sheet1" --range A1:D10
dotnet run -- read-range --path data.xlsx --sheet "Sheet1" --range A1:D10 --format csv
dotnet run -- read-range --path data.xlsx --sheet "Sheet1" --range A1:D10 --format json
```

#### Write Operations

**write-cell** - Write a value to a cell
```bash
dotnet run -- write-cell --path data.xlsx --sheet "Sheet1" --cell A1 --value "Hello"
```

**insert-formula** - Insert an Excel formula
```bash
dotnet run -- insert-formula --path data.xlsx --sheet "Sheet1" --cell C1 --formula "=SUM(A1:B1)"
```

#### Sheet Management

**create-sheet** - Create a new worksheet
```bash
dotnet run -- create-sheet --path data.xlsx --name "NewSheet"
```

**delete-sheet** - Delete a worksheet
```bash
dotnet run -- delete-sheet --path data.xlsx --name "OldSheet"
```

**rename-sheet** - Rename a worksheet
```bash
dotnet run -- rename-sheet --path data.xlsx --old-name "Sheet1" --new-name "Data2025"
```

**copy-sheet** - Copy a worksheet
```bash
dotnet run -- copy-sheet --source data.xlsx --sheet "Sheet1" --target backup.xlsx
dotnet run -- copy-sheet --source data.xlsx --sheet "Template" --target report.xlsx --new-name "January"
```

#### Search and Export

**find-value** - Search for a value in a worksheet
```bash
dotnet run -- find-value --path data.xlsx --sheet "Sheet1" --value "Total"
dotnet run -- find-value --path data.xlsx --sheet "Sheet1" --value "Total" --all
```

**export-sheet** - Export worksheet to CSV or JSON
```bash
dotnet run -- export-sheet --path data.xlsx --sheet "Sheet1" --output output.csv --format csv
dotnet run -- export-sheet --path data.xlsx --sheet "Data" --output data.json --format json
```

**import-data** - Import data from CSV or JSON
```bash
dotnet run -- import-data --path data.xlsx --sheet "Sheet1" --input input.csv --start-cell A1
dotnet run -- import-data --path data.xlsx --sheet "Data" --input data.json --start-cell B2
```

## Development

### Project Structure

```
excel-cli/
├── src/
│   └── ExcelCli/              # Main CLI application
│       ├── Commands/          # Command implementations
│       ├── Services/          # Business logic
│       └── Program.cs         # Entry point
├── tests/
│   └── ExcelCli.Tests/        # Unit tests
├── Plan.md                    # Detailed implementation plan
├── Progress.md                # Progress tracker
└── excel-cli.slnx            # Solution file (XML format)
```

### Building

```bash
dotnet build
```

### Testing

```bash
dotnet test
```

### Code Quality

This project uses SonarAnalyzer.CSharp for code quality analysis. The analyzer is automatically run during build.

## Technologies Used

- **.NET 10.0**: Latest .NET runtime
- **ClosedXML**: Excel file manipulation library
- **System.CommandLine**: Modern CLI framework
- **Serilog**: Structured logging
- **xUnit**: Testing framework
- **NSubstitute**: Mocking framework
- **SonarAnalyzer.CSharp**: Code quality analysis

## Solution Format

This project uses the **slnx** (XML-based) solution format. To migrate an existing .sln file:

```bash
dotnet sln migrate
```

## Contributing

1. Ensure all tests pass
2. Follow existing code style
3. Update Progress.md after each change
4. Run SonarAnalyzer and fix any issues

## License

See LICENSE file for details.

## Project Requirements

- Every project must have the SonarAnalyzer.CSharp NuGet package installed
- Solution uses slnx format for better Git compatibility
- Progress must be tracked in Progress.md after each significant change