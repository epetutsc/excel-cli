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
