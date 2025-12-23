using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Create sheet command
/// </summary>
public class CreateSheetCommand : Command
{
    public CreateSheetCommand(IExcelService excelService, ILogger logger) : base("create-sheet", "Create a new worksheet in an Excel file")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var nameOption = new Option<string>(
            name: "--name",
            description: "Sheet name");
        nameOption.AddAlias("-n");
        nameOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(nameOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var name = context.ParseResult.GetValueForOption(nameOption)!;
            
            try
            {
                await excelService.CreateSheetAsync(path, name);
                Console.WriteLine($"Successfully created sheet '{name}'");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error creating sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Delete sheet command
/// </summary>
public class DeleteSheetCommand : Command
{
    public DeleteSheetCommand(IExcelService excelService, ILogger logger) : base("delete-sheet", "Delete a worksheet from an Excel file")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var nameOption = new Option<string>(
            name: "--name",
            description: "Sheet name");
        nameOption.AddAlias("-n");
        nameOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(nameOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var name = context.ParseResult.GetValueForOption(nameOption)!;
            
            try
            {
                await excelService.DeleteSheetAsync(path, name);
                Console.WriteLine($"Successfully deleted sheet '{name}'");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error deleting sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Rename sheet command
/// </summary>
public class RenameSheetCommand : Command
{
    public RenameSheetCommand(IExcelService excelService, ILogger logger) : base("rename-sheet", "Rename an existing worksheet")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var oldNameOption = new Option<string>(
            name: "--old-name",
            description: "Current sheet name");
        oldNameOption.AddAlias("-o");
        oldNameOption.IsRequired = true;

        var newNameOption = new Option<string>(
            name: "--new-name",
            description: "New sheet name");
        newNameOption.AddAlias("-n");
        newNameOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(oldNameOption);
        AddOption(newNameOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var oldName = context.ParseResult.GetValueForOption(oldNameOption)!;
            var newName = context.ParseResult.GetValueForOption(newNameOption)!;
            
            try
            {
                await excelService.RenameSheetAsync(path, oldName, newName);
                Console.WriteLine($"Successfully renamed sheet '{oldName}' to '{newName}'");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error renaming sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Copy sheet command
/// </summary>
public class CopySheetCommand : Command
{
    public CopySheetCommand(IExcelService excelService, ILogger logger) : base("copy-sheet", "Copy a worksheet within the same file or to another file")
    {
        var sourceOption = new Option<string>(
            name: "--source",
            description: "Source Excel file");
        sourceOption.AddAlias("-s");
        sourceOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Sheet name to copy");
        sheetOption.AddAlias("-sh");
        sheetOption.IsRequired = true;

        var targetOption = new Option<string>(
            name: "--target",
            description: "Target Excel file");
        targetOption.AddAlias("-t");
        targetOption.IsRequired = true;

        var newNameOption = new Option<string?>(
            name: "--new-name",
            description: "New sheet name (optional)");
        newNameOption.AddAlias("-n");

        AddOption(sourceOption);
        AddOption(sheetOption);
        AddOption(targetOption);
        AddOption(newNameOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var source = context.ParseResult.GetValueForOption(sourceOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var target = context.ParseResult.GetValueForOption(targetOption)!;
            var newName = context.ParseResult.GetValueForOption(newNameOption);
            
            try
            {
                await excelService.CopySheetAsync(source, sheet, target, newName);
                var displayName = newName ?? sheet;
                Console.WriteLine($"Successfully copied sheet '{sheet}' to '{target}' as '{displayName}'");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error copying sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Find value command
/// </summary>
public class FindValueCommand : Command
{
    public FindValueCommand(IExcelService excelService, ILogger logger) : base("find-value", "Search for a specific value in a worksheet")
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

        var valueOption = new Option<string>(
            name: "--value",
            description: "Value to search for");
        valueOption.AddAlias("-v");
        valueOption.IsRequired = true;

        var allOption = new Option<bool>(
            name: "--all",
            description: "Find all occurrences",
            getDefaultValue: () => false);
        allOption.AddAlias("-a");

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(valueOption);
        AddOption(allOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var value = context.ParseResult.GetValueForOption(valueOption)!;
            var all = context.ParseResult.GetValueForOption(allOption);
            
            try
            {
                var results = await excelService.FindValueAsync(path, sheet, value, all);
                var resultList = results.ToList();
                
                if (resultList.Count == 0)
                {
                    Console.WriteLine("No matches found.");
                }
                else
                {
                    Console.WriteLine($"Found {resultList.Count} match(es):");
                    foreach (var result in resultList)
                    {
                        Console.WriteLine($"  {result.CellAddress}: {result.Value}");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error finding value");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Insert formula command
/// </summary>
public class InsertFormulaCommand : Command
{
    public InsertFormulaCommand(IExcelService excelService, ILogger logger) : base("insert-formula", "Insert a formula into a cell or range")
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
            description: "Cell address");
        cellOption.AddAlias("-c");
        cellOption.IsRequired = true;

        var formulaOption = new Option<string>(
            name: "--formula",
            description: "Formula (e.g., =SUM(A1:B1))");
        formulaOption.AddAlias("-f");
        formulaOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(cellOption);
        AddOption(formulaOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var cell = context.ParseResult.GetValueForOption(cellOption)!;
            var formula = context.ParseResult.GetValueForOption(formulaOption)!;
            
            try
            {
                await excelService.InsertFormulaAsync(path, sheet, cell, formula);
                Console.WriteLine($"Successfully inserted formula '{formula}' into {cell}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error inserting formula");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Export sheet command
/// </summary>
public class ExportSheetCommand : Command
{
    public ExportSheetCommand(IExcelService excelService, ILogger logger) : base("export-sheet", "Export a worksheet to CSV or JSON format")
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

        var outputOption = new Option<string>(
            name: "--output",
            description: "Output file path");
        outputOption.AddAlias("-o");
        outputOption.IsRequired = true;

        var formatOption = new Option<string>(
            name: "--format",
            description: "Output format (csv, json)");
        formatOption.AddAlias("-f");
        formatOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(outputOption);
        AddOption(formatOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var output = context.ParseResult.GetValueForOption(outputOption)!;
            var format = context.ParseResult.GetValueForOption(formatOption)!;
            
            try
            {
                await excelService.ExportSheetAsync(path, sheet, output, format);
                Console.WriteLine($"Successfully exported sheet '{sheet}' to '{output}' as {format.ToUpper()}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error exporting sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}

/// <summary>
/// Import data command
/// </summary>
public class ImportDataCommand : Command
{
    public ImportDataCommand(IExcelService excelService, ILogger logger) : base("import-data", "Import data from CSV or JSON into a worksheet")
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

        var inputOption = new Option<string>(
            name: "--input",
            description: "Input file path (CSV or JSON)");
        inputOption.AddAlias("-i");
        inputOption.IsRequired = true;

        var startCellOption = new Option<string>(
            name: "--start-cell",
            description: "Starting cell address",
            getDefaultValue: () => "A1");
        startCellOption.AddAlias("-c");

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(inputOption);
        AddOption(startCellOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var input = context.ParseResult.GetValueForOption(inputOption)!;
            var startCell = context.ParseResult.GetValueForOption(startCellOption) ?? "A1";
            
            try
            {
                await excelService.ImportDataAsync(path, sheet, input, startCell);
                Console.WriteLine($"Successfully imported data from '{input}' to sheet '{sheet}' starting at {startCell}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error importing data");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
