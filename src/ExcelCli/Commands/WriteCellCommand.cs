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
