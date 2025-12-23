using System.CommandLine;
using System.CommandLine.Invocation;
using ExcelCli.Services;
using Serilog;

namespace ExcelCli.Commands;

/// <summary>
/// Insert formula command
/// </summary>
public class InsertFormulaCommand : Command
{
    public InsertFormulaCommand(IExcelService excelService, ILogger logger) : base("insert-formula", "Insert a formula into a cell or range")
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

        var cellOption = new Option<string>(
            name: "--cell",
            description: "Cell address");
        cellOption.AddAlias("-c");
        cellOption.IsRequired = true;

        var formulaOption = new Option<string>(
            name: "--formula",
            description: "Formula (e.g., =SUM(A1:B1))");
        formulaOption.AddAlias("-f");
        formulaOption.IsRequired = true;

        AddOption(pathOption);
        AddOption(sheetOption);
        AddOption(cellOption);
        AddOption(formulaOption);

        this.SetHandler(async (InvocationContext context) =>
        {
            var path = context.ParseResult.GetValueForOption(pathOption)!;
            var sheet = context.ParseResult.GetValueForOption(sheetOption)!;
            var cell = context.ParseResult.GetValueForOption(cellOption)!;
            var formula = context.ParseResult.GetValueForOption(formulaOption)!;
            
            try
            {
                await excelService.InsertFormulaAsync(path, sheet, cell, formula);
                Console.WriteLine($"Successfully inserted formula '{formula}' into {cell}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error inserting formula");
                Console.Error.WriteLine($"Error: {ex.Message}");
                context.ExitCode = 1;
            }
        });
    }
}
