using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Rename sheet command
/// </summary>
public class RenameSheetCommand : Command
{
    public RenameSheetCommand(IExcelService excelService, ILogger logger) : base("rename-sheet", "Rename an existing worksheet")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var oldNameOption = new Option<string>(
            name: "--old-name",
            description: "Current sheet name");
        oldNameOption.AddAlias("-o");
        oldNameOption.IsRequired = true;

        var newNameOption = new Option<string>(
            name: "--new-name",
            description: "New sheet name");
        newNameOption.AddAlias("-n");
        newNameOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(oldNameOption);
        AddOption(newNameOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var oldName = context.ParseResult.GetValueForOption(oldNameOption)!;
            var newName = context.ParseResult.GetValueForOption(newNameOption)!;
            
            try
            {
                await excelService.RenameSheetAsync(path, oldName, newName);
                Console.WriteLine($"Successfully renamed sheet '{oldName}' to '{newName}'");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error renaming sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
