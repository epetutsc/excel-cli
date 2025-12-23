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
    public ExportSheetCommand(IExcelService excelService, ILogger logger) : base("export-sheet", 
        "Export a worksheet from an Excel file to CSV or JSON format. " +
        "This is a read-only operation - the Excel file is not modified. " +
        "CSV format: Creates a comma-separated values file with proper escaping for commas and quotes. " +
        "JSON format: Creates a JSON file with data as an array of arrays (rows and columns). " +
        "The output file will be created or overwritten if it exists. " +
        "Examples: excel-cli export-sheet -p data.xlsx -s Sheet1 -o output.csv -f csv | excel-cli export-sheet --path data.xlsx --sheet Data --output data.json --format json")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the source Excel file (.xlsx format). Can be absolute or relative. The file must exist.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Name of the worksheet to export. Must exist in the workbook. Case-sensitive.");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var outputOption = new Option<string>(
            name: "--output",
            description: "Path for the output file. Can be absolute or relative. Will be created or overwritten. Extension should match format (.csv or .json).");
        outputOption.AddAlias("-o");
        outputOption.IsRequired = true;

        var formatOption = new Option<string>(
            name: "--format",
            description: "Output format: 'csv' (comma-separated values) or 'json' (JSON array of arrays). Must be lowercase.");
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
