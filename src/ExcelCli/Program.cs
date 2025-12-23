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
    var rootCommand = new RootCommand("Excel CLI - Tool for reading and modifying Excel files");

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
