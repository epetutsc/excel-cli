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
    public CopySheetCommand(IExcelService excelService, ILogger logger) : base("copy-sheet", "Copy a worksheet within the same file or to another file")
    {
        var sourceOption = new Option<string>(
            name: "--source",
            description: "Source Excel file");
        sourceOption.AddAlias("-s");
        sourceOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Sheet name to copy");
        sheetOption.AddAlias("-sh");
        sheetOption.IsRequired = true;

        var targetOption = new Option<string>(
            name: "--target",
            description: "Target Excel file");
        targetOption.AddAlias("-t");
        targetOption.IsRequired = true;

        var newNameOption = new Option<string?>(
            name: "--new-name",
            description: "New sheet name (optional)");
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
