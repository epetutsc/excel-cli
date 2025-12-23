using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Export sheet command
/// </summary>
public class ExportSheetCommand : Command
{
    public ExportSheetCommand(IExcelService excelService, ILogger logger) : base("export-sheet", "Export a worksheet to CSV or JSON format")
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

        var outputOption = new Option<string>(
            name: "--output",
            description: "Output file path");
        outputOption.AddAlias("-o");
        outputOption.IsRequired = true;

        var formatOption = new Option<string>(
            name: "--format",
            description: "Output format (csv, json)");
        formatOption.AddAlias("-f");
        formatOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(outputOption);
        AddOption(formatOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var output = context.ParseResult.GetValueForOption(outputOption)!;
            var format = context.ParseResult.GetValueForOption(formatOption)!;
            
            try
            {
                await excelService.ExportSheetAsync(path, sheet, output, format);
                Console.WriteLine($"Successfully exported sheet '{sheet}' to '{output}' as {format.ToUpper()}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error exporting sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
