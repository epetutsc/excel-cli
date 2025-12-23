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
    public DeleteSheetCommand(IExcelService excelService, ILogger logger) : base("delete-sheet", 
        "Delete a worksheet from an Excel file. " +
        "This command PERMANENTLY MODIFIES the Excel file by removing the specified worksheet and all its data. " +
        "The file must exist and the sheet name must exist in the workbook. This operation cannot be undone. " +
        "Be cautious when using this command as all data in the deleted sheet will be lost. " +
        "Examples: excel-cli delete-sheet -p data.xlsx -n OldSheet | excel-cli delete-sheet --path data.xlsx --name Temporary")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist and be writable.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var nameOption = new Option<string>(
            name: "--name",
            description: "Name of the worksheet to delete. Must match exactly (case-sensitive). The sheet must exist in the workbook.");
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
