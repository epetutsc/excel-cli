using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Get cell value command - returns the evaluated value from a cell (calculated result if it contains a formula)
/// </summary>
public class GetCellValueCommand : Command
{
    public GetCellValueCommand(IExcelService excelService, ILogger logger) : base("get-cell-value", 
        "Get the evaluated value from a cell. " +
        "If the cell contains a formula, this command returns the calculated result, NOT the formula itself. " +
        "For example, if cell A1 contains '=SUM(B1:B5)', this returns the sum value like '150', not the formula text. " +
        "Cell addresses use Excel's A1 notation where letters represent columns (A, B, C, ..., Z, AA, AB, ...) and numbers represent rows (1, 2, 3, ...). " +
        "This is a read-only operation that does not modify the file. " +
        "Examples: excel-cli get-cell-value --path data.xlsx --sheet Sheet1 --cell A1 | excel-cli get-cell-value -p data.xlsx -s Data -c C5")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Name of the worksheet to read from. Sheet names are case-sensitive and must match exactly.");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var cellOption = new Option<string>(
            name: "--cell",
            description: "Cell address in A1 notation (e.g., A1, B5, Z100, AA1). Column letters are case-insensitive.");
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
                var value = await excelService.GetCellValueAsync(path, sheet, cell);
                Console.WriteLine($"{cell}: {value}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error getting cell value");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
