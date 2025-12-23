# Excel CLI - Terminal Usage

CLI tool for reading and modifying Excel files using ClosedXML. This tool is packaged as a .NET global tool.

## Installation

Install the tool globally:
```bash
dotnet tool install --global ExcelCli
```

Run commands with:
```bash
excel-cli <command> [options]
```

For development, you can also run without installing:
```bash
dotnet run --project src/ExcelCli/ExcelCli.csproj -- <command> [options]
```

## Project Setup Requirements

### Sonar Analyzer
Every project in this solution **MUST** have the SonarAnalyzer.CSharp NuGet package installed for code quality analysis.

```bash
dotnet add package SonarAnalyzer.CSharp
```

### Solution Format
This project uses the **slnx** solution format (XML-based solution file). If you have an existing .sln file, migrate it using:

```bash
dotnet sln migrate
```

### Progress Tracking
After each significant change, progress **MUST** be tracked in the `Progress.md` file. This file should be kept concise to provide a quick overview and allow work to be continued at any time. Update Progress.md after completing each task or milestone.

### Treat Warnings As Errors
Every project **MUST** have `<TreatWarningsAsErrors>true</TreatWarningsAsErrors>` enabled in the `.csproj` file to ensure code quality.

### File Organization
Each class **MUST** be stored in its own file. The file name should match the class name.

### File System Access
For all file system operations, use **System.IO.Abstractions** instead of direct `System.IO` calls. This enables proper testing with mock file systems without creating actual files on disk.

```bash
dotnet add package System.IO.Abstractions
```

In tests, use:
```bash
dotnet add package System.IO.Abstractions.TestingHelpers
```

### Code Coverage
Code coverage **MUST** be at least 90%. Test projects are excluded from code coverage requirements.

## General Information

This tool uses ClosedXML for Excel file operations and supports . xlsx format.  All file paths should be valid and accessible.

## File Support

All commands work with Excel files (. xlsx format). Use absolute or relative paths to specify files.

## Excel Commands

### read-file
Read and display information about an Excel file. 

```bash
excel-cli read-file --path <FILE_PATH>
excel-cli read-file -p data.xlsx
excel-cli read-file --path reports/report.xlsx --sheet "Sheet1"
```

### list-sheets
List all worksheets in an Excel file.

```bash
excel-cli list-sheets --path <FILE_PATH>
excel-cli list-sheets -p data.xlsx
```

### read-cell
Read the content of a specific cell. If the cell contains a formula, returns the formula itself (e.g., "=SUM(A1:B1)"). If the cell contains a value, returns that value.

```bash
excel-cli read-cell --path <FILE_PATH> --sheet <SHEET_NAME> --cell <CELL_ADDRESS>
excel-cli read-cell -p data.xlsx -s "Sheet1" -c A1
excel-cli read-cell --path data.xlsx --sheet "Data" --cell B5
```

### get-cell-value
Get the evaluated value from a cell. If the cell contains a formula, returns the calculated result, NOT the formula. This is useful when you need the actual computed value.

```bash
excel-cli get-cell-value --path <FILE_PATH> --sheet <SHEET_NAME> --cell <CELL_ADDRESS>
excel-cli get-cell-value -p data.xlsx -s "Sheet1" -c C1
excel-cli get-cell-value --path data.xlsx --sheet "Data" --cell D5
```

### read-range
Read a range of cells from a worksheet.

```bash
excel-cli read-range --path <FILE_PATH> --sheet <SHEET_NAME> --range <RANGE>
excel-cli read-range -p data.xlsx -s "Sheet1" -r A1:D10
excel-cli read-range --path data.xlsx --sheet "Data" --range A1:Z100 --format csv
```

### write-cell
Write a value to a specific cell.

```bash
excel-cli write-cell --path <FILE_PATH> --sheet <SHEET_NAME> --cell <CELL_ADDRESS> --value <VALUE>
excel-cli write-cell -p data.xlsx -s "Sheet1" -c A1 -v "Hello"
excel-cli write-cell --path data.xlsx --sheet "Data" --cell B5 --value 42
```

### write-range
Write multiple values to a range of cells. 

```bash
excel-cli write-range --path <FILE_PATH> --sheet <SHEET_NAME> --range <RANGE> --data <DATA>
excel-cli write-range -p data.xlsx -s "Sheet1" -r A1:B2 -d "[[1,2],[3,4]]"
excel-cli write-range --path data.xlsx --sheet "Data" --range A1:C1 --data-file values.json
```

### create-sheet
Create a new worksheet in an Excel file.

```bash
excel-cli create-sheet --path <FILE_PATH> --name <SHEET_NAME>
excel-cli create-sheet -p data.xlsx -n "NewSheet"
excel-cli create-sheet --path data.xlsx --name "Report2025"
```

### delete-sheet
Delete a worksheet from an Excel file.

```bash
excel-cli delete-sheet --path <FILE_PATH> --name <SHEET_NAME>
excel-cli delete-sheet -p data.xlsx -n "OldSheet"
```

### rename-sheet
Rename an existing worksheet.

```bash
excel-cli rename-sheet --path <FILE_PATH> --old-name <OLD_NAME> --new-name <NEW_NAME>
excel-cli rename-sheet -p data.xlsx -o "Sheet1" -n "Data2025"
```

### copy-sheet
Copy a worksheet within the same file or to another file.

```bash
excel-cli copy-sheet --source <SOURCE_FILE> --sheet <SHEET_NAME> --target <TARGET_FILE>
excel-cli copy-sheet -s data.xlsx -sh "Sheet1" -t backup.xlsx
excel-cli copy-sheet --source data.xlsx --sheet "Template" --target report.xlsx --new-name "January"
```

### format-cells
Apply formatting to cells (font, color, borders, etc.).

```bash
excel-cli format-cells --path <FILE_PATH> --sheet <SHEET_NAME> --range <RANGE> --style <STYLE>
excel-cli format-cells -p data. xlsx -s "Sheet1" -r A1:D1 --bold --background-color FF0000
excel-cli format-cells --path data.xlsx --sheet "Data" --range A1:Z1 --font-size 14 --border
```

### find-value
Search for a specific value in a worksheet.

```bash
excel-cli find-value --path <FILE_PATH> --sheet <SHEET_NAME> --value <VALUE>
excel-cli find-value -p data.xlsx -s "Sheet1" -v "Total"
excel-cli find-value --path data.xlsx --sheet "Data" --value 42 --all
```

### export-sheet
Export a worksheet to CSV or JSON format.

```bash
excel-cli export-sheet --path <FILE_PATH> --sheet <SHEET_NAME> --output <OUTPUT_FILE> --format <FORMAT>
excel-cli export-sheet -p data.xlsx -s "Sheet1" -o output.csv -f csv
excel-cli export-sheet --path data.xlsx --sheet "Data" --output data.json --format json
```

### import-data
Import data from CSV or JSON into a worksheet.

```bash
excel-cli import-data --path <FILE_PATH> --sheet <SHEET_NAME> --input <INPUT_FILE> --start-cell <CELL>
excel-cli import-data -p data.xlsx -s "Sheet1" -i input.csv -c A1
excel-cli import-data --path data.xlsx --sheet "Data" --input data.json --start-cell B2
```

### create-table
Create a formatted Excel table from a range. 

```bash
excel-cli create-table --path <FILE_PATH> --sheet <SHEET_NAME> --range <RANGE> --table-name <NAME>
excel-cli create-table -p data.xlsx -s "Sheet1" -r A1:D10 -t "DataTable"
excel-cli create-table --path data.xlsx --sheet "Data" --range A1:Z100 --table-name "SalesData" --style medium
```

### insert-formula
Insert a formula into a cell or range.

```bash
excel-cli insert-formula --path <FILE_PATH> --sheet <SHEET_NAME> --cell <CELL> --formula <FORMULA>
excel-cli insert-formula -p data.xlsx -s "Sheet1" -c C1 -f "=SUM(A1:B1)"
excel-cli insert-formula --path data.xlsx --sheet "Data" --cell D10 --formula "=AVERAGE(D1:D9)"
```

## ClosedXML Best Practices

- Always dispose of workbook objects properly (use `using` statements)
- Use strongly-typed cell values where possible
- Handle errors gracefully (file not found, sheet not exists, etc.)
- Validate cell addresses and ranges before operations
- Use batch operations for better performance when modifying multiple cells
- Consider memory usage for large Excel files
- Always save workbooks after modifications
- Use `IXLRange` for efficient range operations
- Leverage ClosedXML's fluent API for cleaner code
- Handle merged cells and protected sheets appropriately

## Error Handling

The CLI should provide clear error messages for: 
- File not found or inaccessible
- Invalid sheet names
- Invalid cell addresses or ranges
- Permission issues
- Corrupted Excel files
- Invalid data formats
- Formula errors

## Development Guidelines

- Follow SOLID, DRY, and KISS principles
- Use async/await for I/O operations where beneficial
- Implement proper logging with Serilog
- Write unit tests with xUnit and NSubstitute
- Validate all user inputs
- Use dependency injection for better testability
- Implement proper exception handling and user-friendly error messages
