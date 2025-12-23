using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Write cell command
/// </summary>
public class WriteCellCommand : Command
{
    public WriteCellCommand(IExcelService excelService, ILogger logger) : base("write-cell", 
        "Write a value to a specific cell in a worksheet. " +
        "This command MODIFIES the Excel file. The file and sheet must already exist. " +
        "Cell addresses use Excel's A1 notation. The value is written as text/string. " +
        "For formulas, use the 'insert-formula' command instead. " +
        "Examples: excel-cli write-cell -p data.xlsx -s Sheet1 -c A1 -v Hello | excel-cli write-cell --path data.xlsx --sheet Data --cell B5 --value 42")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist and be writable.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Name of the worksheet to write to. Sheet must already exist. Sheet names are case-sensitive.");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var cellOption = new Option<string>(
            name: "--cell",
            description: "Cell address in A1 notation (e.g., A1, B5, Z100). Column letters are case-insensitive.");
        cellOption.AddAlias("-c");
        cellOption.IsRequired = true;

        var valueOption = new Option<string>(
            name: "--value",
            description: "Value to write to the cell. Written as text/string. Use 'insert-formula' command for Excel formulas.");
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
