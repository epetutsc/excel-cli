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
    public ReadRangeCommand(IExcelService excelService, ILogger logger) : base("read-range", "Read a range of cells from a worksheet")
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

        var rangeOption = new Option<string>(
            name: "--range",
            description: "Range (e.g., A1:D10)");
        rangeOption.AddAlias("-r");
        rangeOption.IsRequired = true;

        var formatOption = new Option<string>(
            name: "--format",
            description: "Output format (table, csv, json)",
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
