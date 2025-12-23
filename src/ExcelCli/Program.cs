using System.CommandLine;
using System.IO.Abstractions;
using ExcelCli.Commands;
using ExcelCli.Services;
using Serilog;

// Configure logging
Log.Logger = new LoggerConfiguration()
    .WriteTo.Console()
    .CreateLogger();

try
{
    // Create services
    var fileSystem = new FileSystem();
    var excelService = new ExcelService(Log.Logger, fileSystem);

    // Create root command
    var rootCommand = new RootCommand(
        "Excel CLI - A command-line tool for reading and modifying Excel files (.xlsx format). " +
        "Supports reading file information, managing worksheets, reading/writing cells and ranges, " +
        "searching for values, importing/exporting data (CSV/JSON), and inserting formulas. " +
        "All file paths can be absolute or relative. Excel files must be in .xlsx format (ClosedXML library). " +
        "Use 'excel-cli <command> --help' for detailed information about each command.");

    // Add all commands
    rootCommand.AddCommand(new ReadFileCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new ListSheetsCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new ReadCellCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new ReadRangeCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new WriteCellCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new CreateSheetCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new DeleteSheetCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new RenameSheetCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new CopySheetCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new FindValueCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new InsertFormulaCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new ExportSheetCommand(excelService, Log.Logger));
    rootCommand.AddCommand(new ImportDataCommand(excelService, Log.Logger));

    // Execute command
    return await rootCommand.InvokeAsync(args);
}
catch (Exception ex)
{
    Log.Fatal(ex, "Application terminated unexpectedly");
    return 1;
}
finally
{
    await Log.CloseAndFlushAsync();
}
