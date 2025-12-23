using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Read file command
/// </summary>
public class ReadFileCommand : Command
{
    public ReadFileCommand(IExcelService excelService, ILogger logger) : base("read-file", 
        "Read and display comprehensive information about an Excel file (.xlsx format). " +
        "Shows file metadata (name, size, last modified date), total number of sheets, and details for each sheet (name, row count, column count). " +
        "This is a read-only operation that does not modify the file. " +
        "Example: excel-cli read-file --path data.xlsx")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute (e.g., /home/user/data.xlsx) or relative (e.g., ./data.xlsx, data.xlsx). The file must exist and be readable.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        AddOption(pathOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            try
            {
                var fileInfo = await excelService.ReadFileInfoAsync(path);
                Console.WriteLine($"File: {fileInfo.FileName}");
                Console.WriteLine($"Size: {FormatFileSize(fileInfo.FileSize)}");
                Console.WriteLine($"Last Modified: {fileInfo.LastModified:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"Number of Sheets: {fileInfo.SheetCount}");
                
                var sheets = await excelService.ListSheetsAsync(path);
                Console.WriteLine("\nSheets:");
                foreach (var sheet in sheets)
                {
                    Console.WriteLine($"  - {sheet.Name} ({sheet.RowCount} rows x {sheet.ColumnCount} columns)");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error reading file information");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }

    private static string FormatFileSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB" };
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len /= 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }
}
