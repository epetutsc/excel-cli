using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Read file command
/// </summary>
public class ReadFileCommand : Command
{
    public ReadFileCommand(IExcelService excelService, ILogger logger) : base("read-file", "Read and display information about an Excel file")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        AddOption(pathOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            try
            {
                var fileInfo = await excelService.ReadFileInfoAsync(path);
                Console.WriteLine($"File: {fileInfo.FileName}");
                Console.WriteLine($"Size: {FormatFileSize(fileInfo.FileSize)}");
                Console.WriteLine($"Last Modified: {fileInfo.LastModified:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"Number of Sheets: {fileInfo.SheetCount}");
                
                var sheets = await excelService.ListSheetsAsync(path);
                Console.WriteLine("\nSheets:");
                foreach (var sheet in sheets)
                {
                    Console.WriteLine($"  - {sheet.Name} ({sheet.RowCount} rows x {sheet.ColumnCount} columns)");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error reading file information");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }

    private static string FormatFileSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB" };
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len /= 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }
}

/// <summary>
/// List sheets command
/// </summary>
public class ListSheetsCommand : Command
{
    public ListSheetsCommand(IExcelService excelService, ILogger logger) : base("list-sheets", "List all worksheets in an Excel file")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        AddOption(pathOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            try
            {
                var sheets = await excelService.ListSheetsAsync(path);
                Console.WriteLine("Worksheets:");
                foreach (var sheet in sheets)
                {
                    Console.WriteLine($"  - {sheet.Name} ({sheet.RowCount} rows x {sheet.ColumnCount} columns)");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error listing sheets");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Read cell command
/// </summary>
public class ReadCellCommand : Command
{
    public ReadCellCommand(IExcelService excelService, ILogger logger) : base("read-cell", "Read the value of a specific cell")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Sheet name");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var cellOption = new Option<string>(
            name: "--cell",
            description: "Cell address (e.g., A1)");
        cellOption.AddAlias("-c");
        cellOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(cellOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var cell = context.ParseResult.GetValueForOption(cellOption)!;
            
            try
            {
                var value = await excelService.ReadCellAsync(path, sheet, cell);
                Console.WriteLine($"{cell}: {value}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error reading cell");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Read range command
/// </summary>
public class ReadRangeCommand : Command
{
    public ReadRangeCommand(IExcelService excelService, ILogger logger) : base("read-range", "Read a range of cells from a worksheet")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Sheet name");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var rangeOption = new Option<string>(
            name: "--range",
            description: "Range (e.g., A1:D10)");
        rangeOption.AddAlias("-r");
        rangeOption.IsRequired = true;

        var formatOption = new Option<string>(
            name: "--format",
            description: "Output format (table, csv, json)",
            getDefaultValue: () => "table");
        formatOption.AddAlias("-f");

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(rangeOption);
        AddOption(formatOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var range = context.ParseResult.GetValueForOption(rangeOption)!;
            var format = context.ParseResult.GetValueForOption(formatOption) ?? "table";
            
            try
            {
                var data = await excelService.ReadRangeAsync(path, sheet, range);
                
                if (format.Equals("csv", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var row in data)
                    {
                        Console.WriteLine(string.Join(",", row.Select(EscapeCsvValue)));
                    }
                }
                else if (format.Equals("json", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(data, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
                }
                else
                {
                    // Table format
                    foreach (var row in data)
                    {
                        Console.WriteLine(string.Join(" | ", row));
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error reading range");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }

    private static string EscapeCsvValue(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }
        return value;
    }
}

/// <summary>
/// Write cell command
/// </summary>
public class WriteCellCommand : Command
{
    public WriteCellCommand(IExcelService excelService, ILogger logger) : base("write-cell", "Write a value to a specific cell")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Sheet name");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var cellOption = new Option<string>(
            name: "--cell",
            description: "Cell address (e.g., A1)");
        cellOption.AddAlias("-c");
        cellOption.IsRequired = true;

        var valueOption = new Option<string>(
            name: "--value",
            description: "Value to write");
        valueOption.AddAlias("-v");
        valueOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(cellOption);
        AddOption(valueOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var cell = context.ParseResult.GetValueForOption(cellOption)!;
            var value = context.ParseResult.GetValueForOption(valueOption)!;
            
            try
            {
                await excelService.WriteCellAsync(path, sheet, cell, value);
                Console.WriteLine($"Successfully wrote '{value}' to {cell}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error writing cell");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
