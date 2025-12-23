using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Create sheet command
/// </summary>
public class CreateSheetCommand : Command
{
    public CreateSheetCommand(IExcelService excelService, ILogger logger) : base("create-sheet", 
        "Create a new empty worksheet in an Excel file. " +
        "This command MODIFIES the Excel file by adding a new worksheet. The file must already exist. " +
        "The new sheet name must be unique within the workbook. Sheet names have restrictions (e.g., cannot contain []:*?/\\). " +
        "The new worksheet is created empty with no data. " +
        "Examples: excel-cli create-sheet -p data.xlsx -n NewSheet | excel-cli create-sheet --path data.xlsx --name Report2025")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist and be writable.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var nameOption = new Option<string>(
            name: "--name",
            description: "Name for the new worksheet. Must be unique in the workbook. Cannot contain: [ ] : * ? / \\. Maximum length is typically 31 characters.");
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
                await excelService.CreateSheetAsync(path, name);
                Console.WriteLine($"Successfully created sheet '{name}'");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error creating sheet");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
