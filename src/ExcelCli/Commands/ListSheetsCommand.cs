using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// List sheets command
/// </summary>
public class ListSheetsCommand : Command
{
    public ListSheetsCommand(IExcelService excelService, ILogger logger) : base("list-sheets", 
        "List all worksheets in an Excel file (.xlsx format). " +
        "Displays the name, row count, and column count for each worksheet. " +
        "This is a read-only operation that does not modify the file. " +
        "Example: excel-cli list-sheets --path data.xlsx")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist.");
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
