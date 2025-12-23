using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Read range command
/// </summary>
public class ReadRangeCommand : Command
{
    public ReadRangeCommand(IExcelService excelService, ILogger logger) : base("read-range", 
        "Read and display a range of cells from a worksheet. " +
        "Ranges use Excel's A1 notation with a colon separator (e.g., A1:D10 reads from cell A1 to D10). " +
        "Output can be formatted as table (human-readable), CSV (comma-separated), or JSON (array of arrays). " +
        "This is a read-only operation. " +
        "Examples: excel-cli read-range -p data.xlsx -s Sheet1 -r A1:D10 | excel-cli read-range --path data.xlsx --sheet Data --range A1:Z100 --format csv")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Name of the worksheet to read from. Sheet names are case-sensitive.");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var rangeOption = new Option<string>(
            name: "--range",
            description: "Cell range in A1:B2 notation (e.g., A1:D10, B2:F20). Format is 'TopLeftCell:BottomRightCell'.");
        rangeOption.AddAlias("-r");
        rangeOption.IsRequired = true;

        var formatOption = new Option<string>(
            name: "--format",
            description: "Output format: 'table' (pipe-separated, human-readable), 'csv' (comma-separated values), or 'json' (JSON array). Default: table",
            getDefaultValue: () => "table");
        formatOption.AddAlias("-f");

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(rangeOption);
        AddOption(formatOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var range = context.ParseResult.GetValueForOption(rangeOption)!;
            var format = context.ParseResult.GetValueForOption(formatOption) ?? "table";
            
            try
            {
                var data = await excelService.ReadRangeAsync(path, sheet, range);
                
                if (format.Equals("csv", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var row in data)
                    {
                        Console.WriteLine(string.Join(",", row.Select(EscapeCsvValue)));
                    }
                }
                else if (format.Equals("json", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(data, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
                }
                else
                {
                    // Table format
                    foreach (var row in data)
                    {
                        Console.WriteLine(string.Join(" | ", row));
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error reading range");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }

    private static string EscapeCsvValue(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }
        return value;
    }
}
