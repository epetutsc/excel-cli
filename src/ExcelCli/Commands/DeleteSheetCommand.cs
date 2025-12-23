using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Delete sheet command
/// </summary>
public class DeleteSheetCommand : Command
{
    public DeleteSheetCommand(IExcelService excelService, ILogger logger) : base("delete-sheet", "Delete a worksheet from an Excel file")
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
                await excelService.DeleteSheetAsync(path, name);
                Console.WriteLine($"Successfully deleted sheet '{name}'");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error deleting sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
