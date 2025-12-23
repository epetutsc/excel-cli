# Excel CLI

A powerful command-line tool for reading and modifying Excel files using ClosedXML. This tool processes local Excel files (.xlsx format) and provides various commands for data manipulation.

## Features

- 📖 **Read Operations**: Read file info, list sheets, read cells (formula or value), get evaluated cell values, read ranges
- ✏️ **Write Operations**: Write to cells, insert formulas
- 📋 **Sheet Management**: Create, delete, rename, and copy worksheets
- 🔍 **Search**: Find values in worksheets
- 📤 **Export**: Export sheets to CSV or JSON
- 📥 **Import**: Import data from CSV or JSON files

## Prerequisites

- .NET 10.0 or later

## Installation

### Global Tool Installation

Install the tool globally to use it from anywhere:

```bash
dotnet tool install --global ExcelCli
```

### Local Tool Installation

For project-specific installation, first create a tool manifest if you don't have one:

```bash
dotnet new tool-manifest
```

Then install the tool locally:

```bash
dotnet tool install ExcelCli
```

And run it using:

```bash
dotnet excel-cli <command> [options]
```

### Installing from Source

If you want to install from source:

```bash
git clone https://github.com/epetutsc/excel-cli.git
cd excel-cli
dotnet pack src/ExcelCli/ExcelCli.csproj
dotnet tool install --global --add-source ./nupkg ExcelCli
```

## Usage

Once installed as a global tool, run commands with:

```bash
excel-cli <command> [options]
```

### Available Commands

#### Read Operations

**read-file** - Display file information and sheet summary
```bash
excel-cli read-file --path data.xlsx
```

**list-sheets** - List all worksheets
```bash
excel-cli list-sheets --path data.xlsx
```

**read-cell** - Read a specific cell (returns formula if present, otherwise value)
```bash
excel-cli read-cell --path data.xlsx --sheet "Sheet1" --cell A1
# If cell A1 contains "=SUM(B1:B5)", this returns "=SUM(B1:B5)"
# If cell A1 contains "Hello", this returns "Hello"
```

**get-cell-value** - Get the evaluated value from a cell (calculated result for formulas)
```bash
excel-cli get-cell-value --path data.xlsx --sheet "Sheet1" --cell A1
# If cell A1 contains "=SUM(B1:B5)", this returns the calculated sum (e.g., "150")
# If cell A1 contains "Hello", this returns "Hello"
```

**read-range** - Read a range of cells
```bash
excel-cli read-range --path data.xlsx --sheet "Sheet1" --range A1:D10
excel-cli read-range --path data.xlsx --sheet "Sheet1" --range A1:D10 --format csv
excel-cli read-range --path data.xlsx --sheet "Sheet1" --range A1:D10 --format json
```

#### Write Operations

**write-cell** - Write a value to a cell
```bash
excel-cli write-cell --path data.xlsx --sheet "Sheet1" --cell A1 --value "Hello"
```

**insert-formula** - Insert an Excel formula
```bash
excel-cli insert-formula --path data.xlsx --sheet "Sheet1" --cell C1 --formula "=SUM(A1:B1)"
```

#### Sheet Management

**create-sheet** - Create a new worksheet
```bash
excel-cli create-sheet --path data.xlsx --name "NewSheet"
```

**delete-sheet** - Delete a worksheet
```bash
excel-cli delete-sheet --path data.xlsx --name "OldSheet"
```

**rename-sheet** - Rename a worksheet
```bash
excel-cli rename-sheet --path data.xlsx --old-name "Sheet1" --new-name "Data2025"
```

**copy-sheet** - Copy a worksheet
```bash
excel-cli copy-sheet --source data.xlsx --sheet "Sheet1" --target backup.xlsx
excel-cli copy-sheet --source data.xlsx --sheet "Template" --target report.xlsx --new-name "January"
```

#### Search and Export

**find-value** - Search for a value in a worksheet
```bash
excel-cli find-value --path data.xlsx --sheet "Sheet1" --value "Total"
excel-cli find-value --path data.xlsx --sheet "Sheet1" --value "Total" --all
```

**export-sheet** - Export worksheet to CSV or JSON
```bash
excel-cli export-sheet --path data.xlsx --sheet "Sheet1" --output output.csv --format csv
excel-cli export-sheet --path data.xlsx --sheet "Data" --output data.json --format json
```

**import-data** - Import data from CSV or JSON
```bash
excel-cli import-data --path data.xlsx --sheet "Sheet1" --input input.csv --start-cell A1
excel-cli import-data --path data.xlsx --sheet "Data" --input data.json --start-cell B2
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

### Packing as Tool

To create a NuGet package:

```bash
dotnet pack src/ExcelCli/ExcelCli.csproj
```

The package will be created in the `nupkg` directory.

### Testing During Development

For development and testing, you can run the tool without installing it:

```bash
dotnet run --project src/ExcelCli/ExcelCli.csproj -- <command> [options]
```

Or install it locally from the packed version:

```bash
dotnet tool install --global --add-source ./nupkg ExcelCli
```

To update an existing installation:

```bash
dotnet tool update --global --add-source ./nupkg ExcelCli
```

To uninstall:

```bash
dotnet tool uninstall --global ExcelCli
```

### Running Tests

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