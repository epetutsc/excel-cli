using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Copy sheet command
/// </summary>
public class CopySheetCommand : Command
{
    public CopySheetCommand(IExcelService excelService, ILogger logger) : base("copy-sheet", 
        "Copy a worksheet from one Excel file to another, or within the same file. " +
        "This command MODIFIES the target Excel file by adding a copy of the worksheet. " +
        "The source file and sheet must exist. The target file must exist (or be created by the service). " +
        "If --new-name is not provided, the original sheet name is used (must be unique in target). " +
        "All cell values, formulas, and formatting are copied. " +
        "Examples: excel-cli copy-sheet -s data.xlsx -sh Sheet1 -t backup.xlsx | excel-cli copy-sheet --source data.xlsx --sheet Template --target report.xlsx --new-name January")
    {
        var sourceOption = new Option<string>(
            name: "--source",
            description: "Path to the source Excel file (.xlsx format) containing the worksheet to copy. Can be absolute or relative. Must exist.");
        sourceOption.AddAlias("-s");
        sourceOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Name of the worksheet to copy from the source file. Must exist in source file. Case-sensitive.");
        sheetOption.AddAlias("-sh");
        sheetOption.IsRequired = true;

        var targetOption = new Option<string>(
            name: "--target",
            description: "Path to the target Excel file (.xlsx format) where the worksheet will be copied. Can be the same as source for in-file copy.");
        targetOption.AddAlias("-t");
        targetOption.IsRequired = true;

        var newNameOption = new Option<string?>(
            name: "--new-name",
            description: "Optional: New name for the copied worksheet in the target file. If not specified, uses the original sheet name. Must be unique in target.");
        newNameOption.AddAlias("-n");

        AddOption(sourceOption);
        AddOption(sheetOption);
        AddOption(targetOption);
        AddOption(newNameOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var source = context.ParseResult.GetValueForOption(sourceOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var target = context.ParseResult.GetValueForOption(targetOption)!;
            var newName = context.ParseResult.GetValueForOption(newNameOption);
            
            try
            {
                await excelService.CopySheetAsync(source, sheet, target, newName);
                var displayName = newName ?? sheet;
                Console.WriteLine($"Successfully copied sheet '{sheet}' to '{target}' as '{displayName}'");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error copying sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
