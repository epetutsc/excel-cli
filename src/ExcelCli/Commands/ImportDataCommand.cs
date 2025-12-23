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
