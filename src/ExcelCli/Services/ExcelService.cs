using ClosedXML.Excel;
using Serilog;
using System.Text.Json;

namespace ExcelCli.Services;

/// <summary>
/// Implementation of Excel file operations using ClosedXML
/// </summary>
public class ExcelService : IExcelService
{
    private readonly ILogger _logger;

    public ExcelService(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public async Task<FileInfo> ReadFileInfoAsync(string filePath)
    {
        ValidateFilePath(filePath);
        _logger.Information("Reading file info for {FilePath}", filePath);

        var fileInfo = new System.IO.FileInfo(filePath);
        using var workbook = new XLWorkbook(filePath);
        
        return await Task.FromResult(new FileInfo(
            fileInfo.Name,
            fileInfo.Length,
            fileInfo.LastWriteTime,
            workbook.Worksheets.Count
        ));
    }

    public async Task<IEnumerable<SheetInfo>> ListSheetsAsync(string filePath)
    {
        ValidateFilePath(filePath);

        using var workbook = new XLWorkbook(filePath);
        var sheets = workbook.Worksheets.Select(ws => new SheetInfo(
            ws.Name,
            ws.RangeUsed()?.RowCount() ?? 0,
            ws.RangeUsed()?.ColumnCount() ?? 0
        )).ToList();

        return await Task.FromResult(sheets);
    }

    public async Task<string> ReadCellAsync(string filePath, string sheetName, string cellAddress)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var cell = worksheet.Cell(cellAddress);
        
        return await Task.FromResult(cell.GetValue<string>());
    }

    public async Task<string[][]> ReadRangeAsync(string filePath, string sheetName, string range)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var rangeObj = worksheet.Range(range);
        
        var data = new List<string[]>();
        foreach (var row in rangeObj.Rows())
        {
            var rowData = row.Cells().Select(c => c.GetValue<string>()).ToArray();
            data.Add(rowData);
        }
        
        return await Task.FromResult(data.ToArray());
    }

    public async Task WriteCellAsync(string filePath, string sheetName, string cellAddress, string value)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        worksheet.Cell(cellAddress).Value = value;
        workbook.Save();
        
        await Task.CompletedTask;
    }

    public async Task WriteRangeAsync(string filePath, string sheetName, string range, string[][] data)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var rangeObj = worksheet.Range(range);
        
        var rowIndex = 0;
        foreach (var row in rangeObj.Rows())
        {
            if (rowIndex >= data.Length) break;
            
            var colIndex = 0;
            foreach (var cell in row.Cells())
            {
                if (colIndex >= data[rowIndex].Length) break;
                cell.Value = data[rowIndex][colIndex];
                colIndex++;
            }
            rowIndex++;
        }
        
        workbook.Save();
        await Task.CompletedTask;
    }

    public async Task CreateSheetAsync(string filePath, string sheetName)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        
        if (workbook.Worksheets.Contains(sheetName))
        {
            throw new InvalidOperationException($"Sheet '{sheetName}' already exists.");
        }
        
        workbook.Worksheets.Add(sheetName);
        workbook.Save();
        
        await Task.CompletedTask;
    }

    public async Task DeleteSheetAsync(string filePath, string sheetName)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        
        if (workbook.Worksheets.Count <= 1)
        {
            throw new InvalidOperationException("Cannot delete the last worksheet in the workbook.");
        }
        
        var worksheet = GetWorksheet(workbook, sheetName);
        worksheet.Delete();
        workbook.Save();
        
        await Task.CompletedTask;
    }

    public async Task RenameSheetAsync(string filePath, string oldName, string newName)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, oldName);
        
        if (workbook.Worksheets.Contains(newName))
        {
            throw new InvalidOperationException($"Sheet '{newName}' already exists.");
        }
        
        worksheet.Name = newName;
        workbook.Save();
        
        await Task.CompletedTask;
    }

    public async Task CopySheetAsync(string sourceFile, string sheetName, string targetFile, string? newName = null)
    {
        ValidateFilePath(sourceFile);
        
        using var sourceWorkbook = new XLWorkbook(sourceFile);
        var sourceWorksheet = GetWorksheet(sourceWorkbook, sheetName);
        
        XLWorkbook targetWorkbook;
        var targetExists = File.Exists(targetFile);
        
        if (targetExists)
        {
            targetWorkbook = new XLWorkbook(targetFile);
        }
        else
        {
            targetWorkbook = new XLWorkbook();
        }
        
        using (targetWorkbook)
        {
            var copiedSheet = sourceWorksheet.CopyTo(targetWorkbook);
            if (!string.IsNullOrEmpty(newName))
            {
                copiedSheet.Name = newName;
            }
            
            targetWorkbook.SaveAs(targetFile);
        }
        
        await Task.CompletedTask;
    }

    public async Task<IEnumerable<CellLocation>> FindValueAsync(string filePath, string sheetName, string value, bool findAll = false)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var results = new List<CellLocation>();
        
        var rangeUsed = worksheet.RangeUsed();
        if (rangeUsed == null)
        {
            return await Task.FromResult(results);
        }
        
        foreach (var cell in rangeUsed.CellsUsed())
        {
            var cellValue = cell.GetValue<string>();
            if (cellValue.Contains(value, StringComparison.OrdinalIgnoreCase))
            {
                var cellAddress = cell.Address.ToString();
                if (!string.IsNullOrEmpty(cellAddress))
                {
                    results.Add(new CellLocation(sheetName, cellAddress, cellValue));
                    
                    if (!findAll)
                    {
                        break;
                    }
                }
            }
        }
        
        return await Task.FromResult(results);
    }

    public async Task ExportSheetAsync(string filePath, string sheetName, string outputFile, string format)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var rangeUsed = worksheet.RangeUsed();
        
        if (rangeUsed == null)
        {
            throw new InvalidOperationException("Worksheet is empty.");
        }
        
        if (format.Equals("csv", StringComparison.OrdinalIgnoreCase))
        {
            await ExportToCsvAsync(rangeUsed, outputFile);
        }
        else if (format.Equals("json", StringComparison.OrdinalIgnoreCase))
        {
            await ExportToJsonAsync(rangeUsed, outputFile);
        }
        else
        {
            throw new ArgumentException($"Unsupported format: {format}");
        }
    }

    public async Task ImportDataAsync(string filePath, string sheetName, string inputFile, string startCell)
    {
        ValidateFilePath(filePath);
        
        if (!File.Exists(inputFile))
        {
            throw new FileNotFoundException($"Input file not found: {inputFile}");
        }
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        
        var extension = Path.GetExtension(inputFile).ToLowerInvariant();
        
        if (extension == ".csv")
        {
            await ImportFromCsvAsync(worksheet, inputFile, startCell);
        }
        else if (extension == ".json")
        {
            await ImportFromJsonAsync(worksheet, inputFile, startCell);
        }
        else
        {
            throw new ArgumentException($"Unsupported file type: {extension}");
        }
        
        workbook.Save();
    }

    public async Task InsertFormulaAsync(string filePath, string sheetName, string cellAddress, string formula)
    {
        ValidateFilePath(filePath);
        
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var cell = worksheet.Cell(cellAddress);
        
        if (!formula.StartsWith('='))
        {
            formula = "=" + formula;
        }
        
        cell.FormulaA1 = formula;
        workbook.Save();
        
        await Task.CompletedTask;
    }

    private static void ValidateFilePath(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }
        
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"File not found: {filePath}");
        }
    }

    private static IXLWorksheet GetWorksheet(XLWorkbook workbook, string sheetName)
    {
        if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var worksheet))
        {
            throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
        }
        
        return worksheet;
    }

    private static async Task ExportToCsvAsync(IXLRange range, string outputFile)
    {
        var lines = new List<string>();
        
        foreach (var row in range.Rows())
        {
            var values = row.Cells().Select(c => EscapeCsvValue(c.GetValue<string>()));
            lines.Add(string.Join(",", values));
        }
        
        await File.WriteAllLinesAsync(outputFile, lines);
    }

    private static string EscapeCsvValue(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }
        return value;
    }

    private static async Task ExportToJsonAsync(IXLRange range, string outputFile)
    {
        var data = new List<Dictionary<string, string>>();
        var headers = range.FirstRow().Cells().Select(c => c.GetValue<string>()).ToArray();
        
        foreach (var row in range.Rows().Skip(1))
        {
            var rowData = new Dictionary<string, string>();
            var cells = row.Cells().ToArray();
            
            for (int i = 0; i < headers.Length && i < cells.Length; i++)
            {
                rowData[headers[i]] = cells[i].GetValue<string>();
            }
            
            data.Add(rowData);
        }
        
        var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(outputFile, json);
    }

    private static async Task ImportFromCsvAsync(IXLWorksheet worksheet, string inputFile, string startCell)
    {
        var lines = await File.ReadAllLinesAsync(inputFile);
        var cell = worksheet.Cell(startCell);
        var startRow = cell.Address.RowNumber;
        var startCol = cell.Address.ColumnNumber;
        
        for (int i = 0; i < lines.Length; i++)
        {
            var values = ParseCsvLine(lines[i]);
            for (int j = 0; j < values.Length; j++)
            {
                worksheet.Cell(startRow + i, startCol + j).Value = values[j];
            }
        }
    }

    private static string[] ParseCsvLine(string line)
    {
        var values = new List<string>();
        var currentValue = new System.Text.StringBuilder();
        var inQuotes = false;
        
#pragma warning disable S127 // Loop counter update is intentional for CSV parsing efficiency
        for (int i = 0; i < line.Length;)
        {
            var c = line[i];
            
            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    // Handle escaped quotes
                    currentValue.Append('"');
                    i += 2; // Skip both quotes
                }
                else
                {
                    inQuotes = !inQuotes;
                    i++;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                values.Add(currentValue.ToString());
                currentValue.Clear();
                i++;
            }
            else
            {
                currentValue.Append(c);
                i++;
            }
        }
#pragma warning restore S127
        
        values.Add(currentValue.ToString());
        return values.ToArray();
    }

    private static async Task ImportFromJsonAsync(IXLWorksheet worksheet, string inputFile, string startCell)
    {
        var json = await File.ReadAllTextAsync(inputFile);
        var data = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(json);
        
        if (data == null || data.Count == 0)
        {
            return;
        }
        
        var cell = worksheet.Cell(startCell);
        var startRow = cell.Address.RowNumber;
        var startCol = cell.Address.ColumnNumber;
        
        // Write headers
        var headers = data[0].Keys.ToArray();
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cell(startRow, startCol + i).Value = headers[i];
        }
        
        // Write data
        for (int i = 0; i < data.Count; i++)
        {
            var row = data[i];
            for (int j = 0; j < headers.Length; j++)
            {
                if (row.TryGetValue(headers[j], out var value))
                {
                    worksheet.Cell(startRow + i + 1, startCol + j).Value = value;
                }
            }
        }
    }
}
