using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Read cell command
/// </summary>
public class ReadCellCommand : Command
{
    public ReadCellCommand(IExcelService excelService, ILogger logger) : base("read-cell", 
        "Read and display the content of a specific cell from a worksheet. " +
        "If the cell contains a formula, this command returns the formula itself (e.g., '=SUM(A1:B1)'), NOT the calculated value. " +
        "If the cell contains a plain value, that value is returned. " +
        "To get the evaluated/calculated result of a formula, use the 'read-cell-value' command instead. " +
        "Cell addresses use Excel's A1 notation where letters represent columns (A, B, C, ..., Z, AA, AB, ...) and numbers represent rows (1, 2, 3, ...). " +
        "This is a read-only operation that does not modify the file. " +
        "Examples: excel-cli read-cell --path data.xlsx --sheet Sheet1 --cell A1 | excel-cli read-cell -p data.xlsx -s Data -c B5")
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
