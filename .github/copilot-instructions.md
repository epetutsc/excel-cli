# Excel CLI - Terminal Usage

CLI tool for reading and modifying Excel files using ClosedXML.  Run commands with `dotnet run -- <command>`.

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
dotnet run -- read-file --path <FILE_PATH>
dotnet run -- read-file -p data.xlsx
dotnet run -- read-file --path reports/report.xlsx --sheet "Sheet1"
```

### list-sheets
List all worksheets in an Excel file.

```bash
dotnet run -- list-sheets --path <FILE_PATH>
dotnet run -- list-sheets -p data.xlsx
```

### read-cell
Read the value of a specific cell.

```bash
dotnet run -- read-cell --path <FILE_PATH> --sheet <SHEET_NAME> --cell <CELL_ADDRESS>
dotnet run -- read-cell -p data.xlsx -s "Sheet1" -c A1
dotnet run -- read-cell --path data.xlsx --sheet "Data" --cell B5
```

### read-range
Read a range of cells from a worksheet.

```bash
dotnet run -- read-range --path <FILE_PATH> --sheet <SHEET_NAME> --range <RANGE>
dotnet run -- read-range -p data.xlsx -s "Sheet1" -r A1:D10
dotnet run -- read-range --path data.xlsx --sheet "Data" --range A1:Z100 --format csv
```

### write-cell
Write a value to a specific cell.

```bash
dotnet run -- write-cell --path <FILE_PATH> --sheet <SHEET_NAME> --cell <CELL_ADDRESS> --value <VALUE>
dotnet run -- write-cell -p data.xlsx -s "Sheet1" -c A1 -v "Hello"
dotnet run -- write-cell --path data.xlsx --sheet "Data" --cell B5 --value 42
```

### write-range
Write multiple values to a range of cells. 

```bash
dotnet run -- write-range --path <FILE_PATH> --sheet <SHEET_NAME> --range <RANGE> --data <DATA>
dotnet run -- write-range -p data.xlsx -s "Sheet1" -r A1:B2 -d "[[1,2],[3,4]]"
dotnet run -- write-range --path data.xlsx --sheet "Data" --range A1:C1 --data-file values.json
```

### create-sheet
Create a new worksheet in an Excel file.

```bash
dotnet run -- create-sheet --path <FILE_PATH> --name <SHEET_NAME>
dotnet run -- create-sheet -p data.xlsx -n "NewSheet"
dotnet run -- create-sheet --path data.xlsx --name "Report2025"
```

### delete-sheet
Delete a worksheet from an Excel file.

```bash
dotnet run -- delete-sheet --path <FILE_PATH> --name <SHEET_NAME>
dotnet run -- delete-sheet -p data.xlsx -n "OldSheet"
```

### rename-sheet
Rename an existing worksheet.

```bash
dotnet run -- rename-sheet --path <FILE_PATH> --old-name <OLD_NAME> --new-name <NEW_NAME>
dotnet run -- rename-sheet -p data.xlsx -o "Sheet1" -n "Data2025"
```

### copy-sheet
Copy a worksheet within the same file or to another file.

```bash
dotnet run -- copy-sheet --source <SOURCE_FILE> --sheet <SHEET_NAME> --target <TARGET_FILE>
dotnet run -- copy-sheet -s data.xlsx -sh "Sheet1" -t backup.xlsx
dotnet run -- copy-sheet --source data.xlsx --sheet "Template" --target report.xlsx --new-name "January"
```

### format-cells
Apply formatting to cells (font, color, borders, etc.).

```bash
dotnet run -- format-cells --path <FILE_PATH> --sheet <SHEET_NAME> --range <RANGE> --style <STYLE>
dotnet run -- format-cells -p data. xlsx -s "Sheet1" -r A1:D1 --bold --background-color FF0000
dotnet run -- format-cells --path data.xlsx --sheet "Data" --range A1:Z1 --font-size 14 --border
```

### find-value
Search for a specific value in a worksheet.

```bash
dotnet run -- find-value --path <FILE_PATH> --sheet <SHEET_NAME> --value <VALUE>
dotnet run -- find-value -p data.xlsx -s "Sheet1" -v "Total"
dotnet run -- find-value --path data.xlsx --sheet "Data" --value 42 --all
```

### export-sheet
Export a worksheet to CSV or JSON format.

```bash
dotnet run -- export-sheet --path <FILE_PATH> --sheet <SHEET_NAME> --output <OUTPUT_FILE> --format <FORMAT>
dotnet run -- export-sheet -p data.xlsx -s "Sheet1" -o output.csv -f csv
dotnet run -- export-sheet --path data.xlsx --sheet "Data" --output data.json --format json
```

### import-data
Import data from CSV or JSON into a worksheet.

```bash
dotnet run -- import-data --path <FILE_PATH> --sheet <SHEET_NAME> --input <INPUT_FILE> --start-cell <CELL>
dotnet run -- import-data -p data.xlsx -s "Sheet1" -i input.csv -c A1
dotnet run -- import-data --path data.xlsx --sheet "Data" --input data.json --start-cell B2
```

### create-table
Create a formatted Excel table from a range. 

```bash
dotnet run -- create-table --path <FILE_PATH> --sheet <SHEET_NAME> --range <RANGE> --table-name <NAME>
dotnet run -- create-table -p data.xlsx -s "Sheet1" -r A1:D10 -t "DataTable"
dotnet run -- create-table --path data.xlsx --sheet "Data" --range A1:Z100 --table-name "SalesData" --style medium
```

### insert-formula
Insert a formula into a cell or range.

```bash
dotnet run -- insert-formula --path <FILE_PATH> --sheet <SHEET_NAME> --cell <CELL> --formula <FORMULA>
dotnet run -- insert-formula -p data.xlsx -s "Sheet1" -c C1 -f "=SUM(A1:B1)"
dotnet run -- insert-formula --path data.xlsx --sheet "Data" --cell D10 --formula "=AVERAGE(D1:D9)"
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
