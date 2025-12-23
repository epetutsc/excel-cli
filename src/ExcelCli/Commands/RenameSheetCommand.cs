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
    public RenameSheetCommand(IExcelService excelService, ILogger logger) : base("rename-sheet", 
        "Rename an existing worksheet in an Excel file. " +
        "This command MODIFIES the Excel file by changing the name of a worksheet. " +
        "The old sheet name must exist, and the new name must be unique in the workbook. " +
        "Sheet names cannot contain: [ ] : * ? / \\. Maximum length is typically 31 characters. " +
        "Examples: excel-cli rename-sheet -p data.xlsx -o Sheet1 -n Data2025 | excel-cli rename-sheet --path data.xlsx --old-name OldName --new-name NewName")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist and be writable.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var oldNameOption = new Option<string>(
            name: "--old-name",
            description: "Current name of the worksheet to rename. Must match exactly (case-sensitive) and must exist in the workbook.");
        oldNameOption.AddAlias("-o");
        oldNameOption.IsRequired = true;

        var newNameOption = new Option<string>(
            name: "--new-name",
            description: "New name for the worksheet. Must be unique in the workbook. Cannot contain: [ ] : * ? / \\. Max 31 characters.");
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
