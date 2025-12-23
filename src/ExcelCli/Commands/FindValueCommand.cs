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
