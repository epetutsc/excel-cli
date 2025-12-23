using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Import data command
/// </summary>
public class ImportDataCommand : Command
{
    public ImportDataCommand(IExcelService excelService, ILogger logger) : base("import-data", 
        "Import data from a CSV or JSON file into an Excel worksheet. " +
        "This command MODIFIES the Excel file by writing data from the input file. " +
        "CSV format: Reads comma-separated values. First row typically contains headers. " +
        "JSON format: Expects an array of arrays structure [[row1], [row2], ...]. " +
        "Data is written starting at the specified cell address (default: A1). " +
        "The Excel file and worksheet must already exist. " +
        "Examples: excel-cli import-data -p data.xlsx -s Sheet1 -i input.csv -c A1 | excel-cli import-data --path data.xlsx --sheet Data --input data.json --start-cell B2")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the target Excel file (.xlsx format). Can be absolute or relative. The file must exist and be writable.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Name of the worksheet where data will be imported. Must exist in the workbook. Case-sensitive.");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var inputOption = new Option<string>(
            name: "--input",
            description: "Path to the input data file (CSV or JSON format). Can be absolute or relative. File format is auto-detected from content/extension.");
        inputOption.AddAlias("-i");
        inputOption.IsRequired = true;

        var startCellOption = new Option<string>(
            name: "--start-cell",
            description: "Cell address in A1 notation where data import begins (top-left corner). Default: A1. Examples: A1, B2, C10",
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
