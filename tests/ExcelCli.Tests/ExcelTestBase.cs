using ClosedXML.Excel;
using ExcelCli.Services;
using NSubstitute;
using Serilog;
using System.IO.Abstractions;
using System.IO.Abstractions.TestingHelpers;

namespace ExcelCli.Tests;

/// <summary>
/// Base class for Excel service tests providing common test setup and utilities
/// </summary>
public abstract class ExcelTestBase : IDisposable
{
    protected readonly ILogger Logger;
    protected readonly MockFileSystem FileSystem;
    protected readonly string TestDirectory;
    private bool _disposed;

    protected ExcelTestBase()
    {
        Logger = Substitute.For<ILogger>();
        FileSystem = new MockFileSystem();
        TestDirectory = Path.Combine(Path.GetTempPath(), $"ExcelCliTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(TestDirectory);
    }

    /// <summary>
    /// Creates a simple test Excel file with specified number of sheets
    /// </summary>
    protected string CreateTestExcelFile(string fileName = "test.xlsx", int sheetCount = 1)
    {
        var filePath = Path.Combine(TestDirectory, fileName);
        using var workbook = new XLWorkbook();
        
        for (int i = 1; i <= sheetCount; i++)
        {
            var sheet = workbook.Worksheets.Add($"Sheet{i}");
            sheet.Cell("A1").Value = $"Sheet{i} Data";
        }
        
        workbook.SaveAs(filePath);
        
        // Register the file in MockFileSystem
        FileSystem.AddFile(filePath, new MockFileData(File.ReadAllBytes(filePath)));
        
        return filePath;
    }

    /// <summary>
    /// Creates a test Excel file with data in a specified range
    /// </summary>
    protected string CreateTestExcelFileWithData(string fileName, string sheetName, string[][] data)
    {
        var filePath = Path.Combine(TestDirectory, fileName);
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add(sheetName);
        
        for (int row = 0; row < data.Length; row++)
        {
            for (int col = 0; col < data[row].Length; col++)
            {
                sheet.Cell(row + 1, col + 1).Value = data[row][col];
            }
        }
        
        workbook.SaveAs(filePath);
        FileSystem.AddFile(filePath, new MockFileData(File.ReadAllBytes(filePath)));
        
        return filePath;
    }

    /// <summary>
    /// Creates a test Excel file with formulas
    /// </summary>
    protected string CreateTestExcelFileWithFormulas(string fileName, string sheetName)
    {
        var filePath = Path.Combine(TestDirectory, fileName);
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add(sheetName);
        
        // Set up values for formulas to reference
        sheet.Cell("A1").Value = 10;
        sheet.Cell("A2").Value = 20;
        sheet.Cell("A3").Value = 30;
        sheet.Cell("B1").Value = 5;
        sheet.Cell("B2").Value = 15;
        sheet.Cell("B3").Value = 25;
        
        // Add formulas
        sheet.Cell("C1").FormulaA1 = "=A1+B1";      // Should be 15
        sheet.Cell("C2").FormulaA1 = "=A2*B2";      // Should be 300
        sheet.Cell("C3").FormulaA1 = "=SUM(A1:A3)"; // Should be 60
        sheet.Cell("D1").FormulaA1 = "=AVERAGE(B1:B3)"; // Should be 15
        
        workbook.SaveAs(filePath);
        FileSystem.AddFile(filePath, new MockFileData(File.ReadAllBytes(filePath)));
        
        return filePath;
    }

    /// <summary>
    /// Creates an ExcelService instance with the test configuration
    /// </summary>
    protected ExcelService CreateService()
    {
        return new ExcelService(Logger, FileSystem);
    }

    /// <summary>
    /// Refreshes file in MockFileSystem after modification
    /// </summary>
    protected void RefreshMockFile(string filePath)
    {
        FileSystem.AddFile(filePath, new MockFileData(File.ReadAllBytes(filePath)));
    }

    /// <summary>
    /// Creates a test text file (CSV, JSON, etc.)
    /// </summary>
    protected string CreateTestTextFile(string fileName, string content)
    {
        var filePath = Path.Combine(TestDirectory, fileName);
        File.WriteAllText(filePath, content);
        FileSystem.AddFile(filePath, new MockFileData(content));
        return filePath;
    }

    /// <summary>
    /// Creates a test CSV file
    /// </summary>
    protected string CreateTestCsvFile(string fileName, string content)
    {
        return CreateTestTextFile(fileName, content);
    }

    /// <summary>
    /// Creates a test JSON file
    /// </summary>
    protected string CreateTestJsonFile(string fileName, string content)
    {
        return CreateTestTextFile(fileName, content);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed && disposing && Directory.Exists(TestDirectory))
        {
            try
            {
                Directory.Delete(TestDirectory, true);
            }
            catch
            {
                // Ignore cleanup errors in tests
            }
            _disposed = true;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
