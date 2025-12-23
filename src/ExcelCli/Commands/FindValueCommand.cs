using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Find value command
/// </summary>
public class FindValueCommand : Command
{
    public FindValueCommand(IExcelService excelService, ILogger logger) : base("find-value", 
        "Search for a specific value in a worksheet and return matching cell addresses. " +
        "By default, returns only the first match. Use --all flag to find all occurrences. " +
        "Search is performed across all cells in the worksheet. This is a read-only operation. " +
        "Returns cell addresses in A1 notation along with the cell values. " +
        "Examples: excel-cli find-value -p data.xlsx -s Sheet1 -v Total | excel-cli find-value --path data.xlsx --sheet Data --value 42 --all")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Name of the worksheet to search in. Sheet names are case-sensitive and must exist in the workbook.");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var valueOption = new Option<string>(
            name: "--value",
            description: "Value to search for in the worksheet. Searches for exact matches in cell values.");
        valueOption.AddAlias("-v");
        valueOption.IsRequired = true;

        var allOption = new Option<bool>(
            name: "--all",
            description: "Find all occurrences of the value. If false (default), returns only the first match. If true, returns all matches.",
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
