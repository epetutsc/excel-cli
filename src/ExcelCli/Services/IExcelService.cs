namespace ExcelCli.Services;

/// <summary>
/// Service for Excel file operations
/// </summary>
public interface IExcelService
{
    /// <summary>
    /// Read information about an Excel file
    /// </summary>
    Task<FileInfo> ReadFileInfoAsync(string filePath);

    /// <summary>
    /// List all worksheets in an Excel file
    /// </summary>
    Task<IEnumerable<SheetInfo>> ListSheetsAsync(string filePath);

    /// <summary>
    /// Read a specific cell value
    /// </summary>
    Task<string> ReadCellAsync(string filePath, string sheetName, string cellAddress);

    /// <summary>
    /// Read a range of cells
    /// </summary>
    Task<string[][]> ReadRangeAsync(string filePath, string sheetName, string range);

    /// <summary>
    /// Write a value to a specific cell
    /// </summary>
    Task WriteCellAsync(string filePath, string sheetName, string cellAddress, string value);

    /// <summary>
    /// Write data to a range of cells
    /// </summary>
#pragma warning disable S2368 // Public methods should not have multidimensional array parameters - Required for Excel range data
    Task WriteRangeAsync(string filePath, string sheetName, string range, string[][] data);
#pragma warning restore S2368

    /// <summary>
    /// Create a new worksheet
    /// </summary>
    Task CreateSheetAsync(string filePath, string sheetName);

    /// <summary>
    /// Delete a worksheet
    /// </summary>
    Task DeleteSheetAsync(string filePath, string sheetName);

    /// <summary>
    /// Rename a worksheet
    /// </summary>
    Task RenameSheetAsync(string filePath, string oldName, string newName);

    /// <summary>
    /// Copy a worksheet
    /// </summary>
    Task CopySheetAsync(string sourceFile, string sheetName, string targetFile, string? newName = null);

    /// <summary>
    /// Find a value in a worksheet
    /// </summary>
    Task<IEnumerable<CellLocation>> FindValueAsync(string filePath, string sheetName, string value, bool findAll = false);

    /// <summary>
    /// Export a worksheet to CSV or JSON
    /// </summary>
    Task ExportSheetAsync(string filePath, string sheetName, string outputFile, string format);

    /// <summary>
    /// Import data from CSV or JSON
    /// </summary>
    Task ImportDataAsync(string filePath, string sheetName, string inputFile, string startCell);

    /// <summary>
    /// Insert a formula into a cell
    /// </summary>
    Task InsertFormulaAsync(string filePath, string sheetName, string cellAddress, string formula);

    /// <summary>
    /// Get the evaluated value from a cell (if cell contains a formula, returns the calculated result)
    /// </summary>
    Task<string> GetCellValueAsync(string filePath, string sheetName, string cellAddress);
}

