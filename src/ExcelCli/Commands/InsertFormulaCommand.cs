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
    public InsertFormulaCommand(IExcelService excelService, ILogger logger) : base("insert-formula", 
        "Insert an Excel formula into a specific cell. " +
        "This command MODIFIES the Excel file by writing a formula to a cell. " +
        "Formulas must start with '=' (e.g., =SUM(A1:B1), =AVERAGE(A1:A10), =A1+B1). " +
        "The formula will be evaluated by Excel when the file is opened. " +
        "Cell references in formulas use A1 notation. Common functions: SUM, AVERAGE, COUNT, IF, VLOOKUP, etc. " +
        "Examples: excel-cli insert-formula -p data.xlsx -s Sheet1 -c C1 -f '=SUM(A1:B1)' | excel-cli insert-formula --path data.xlsx --sheet Data --cell D10 --formula '=AVERAGE(D1:D9)'")
    {
        var pathOption = new Option<string>(
            name: "--path",
            description: "Path to the Excel file (.xlsx format). Can be absolute or relative. The file must exist and be writable.");
        pathOption.AddAlias("-p");
        pathOption.IsRequired = true;

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Name of the worksheet where the formula will be inserted. Must exist in the workbook. Case-sensitive.");
        sheetOption.AddAlias("-s");
        sheetOption.IsRequired = true;

        var cellOption = new Option<string>(
            name: "--cell",
            description: "Cell address in A1 notation where the formula will be inserted (e.g., A1, C5, Z100).");
        cellOption.AddAlias("-c");
        cellOption.IsRequired = true;

        var formulaOption = new Option<string>(
            name: "--formula",
            description: "Excel formula to insert. Must start with '='. Examples: =SUM(A1:B1), =AVERAGE(A1:A10), =IF(A1>10,\"Yes\",\"No\"), =A1*B1");
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
