using ClosedXML.Excel;
using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for ImportDataAsync operation
/// </summary>
public class ImportDataTests : ExcelTestBase
{
    [Fact]
    public async Task ImportDataAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();
        var inputPath = CreateTestCsvFile("input.csv", "A,B\n1,2");

        await Assert.ThrowsAsync<ArgumentException>(() => service.ImportDataAsync(null!, "Sheet1", inputPath, "A1"));
    }

    [Fact]
    public async Task ImportDataAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";
        var inputPath = CreateTestCsvFile("input_test.csv", "A,B\n1,2");

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ImportDataAsync(nonExistentPath, "Sheet1", inputPath, "A1"));
    }

    [Fact]
    public async Task ImportDataAsync_WithNonExistentInputFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("import_no_input.xlsx", 1);

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ImportDataAsync(filePath, "Sheet1", "/tmp/nonexistent.csv", "A1"));
    }

    [Fact]
    public async Task ImportDataAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("import_no_sheet.xlsx", 1);
        var inputPath = CreateTestCsvFile("import_input.csv", "A,B\n1,2");

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.ImportDataAsync(filePath, "NonExistent", inputPath, "A1"));
    }

    [Fact]
    public async Task ImportDataAsync_FromCsv_ImportsData()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("import_csv.xlsx", 1);
        var csvContent = "Name,Age\nAlice,30\nBob,25";
        var inputPath = CreateTestCsvFile("import.csv", csvContent);

        await service.ImportDataAsync(filePath, "Sheet1", inputPath, "A1");

        using var workbook = new XLWorkbook(filePath);
        var sheet = workbook.Worksheet("Sheet1");
        Assert.Equal("Name", sheet.Cell("A1").GetValue<string>());
        Assert.Equal("Age", sheet.Cell("B1").GetValue<string>());
        Assert.Equal("Alice", sheet.Cell("A2").GetValue<string>());
        Assert.Equal("30", sheet.Cell("B2").GetValue<string>());
    }

    [Fact]
    public async Task ImportDataAsync_FromJson_ImportsData()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("import_json.xlsx", 1);
        var jsonContent = "[{\"Name\":\"Charlie\",\"Score\":\"95\"},{\"Name\":\"Diana\",\"Score\":\"88\"}]";
        var inputPath = CreateTestJsonFile("import.json", jsonContent);

        await service.ImportDataAsync(filePath, "Sheet1", inputPath, "A1");

        using var workbook = new XLWorkbook(filePath);
        var sheet = workbook.Worksheet("Sheet1");
        // JSON import adds headers
        Assert.Equal("Name", sheet.Cell("A1").GetValue<string>());
        Assert.Equal("Score", sheet.Cell("B1").GetValue<string>());
        Assert.Equal("Charlie", sheet.Cell("A2").GetValue<string>());
        Assert.Equal("95", sheet.Cell("B2").GetValue<string>());
    }

    [Fact]
    public async Task ImportDataAsync_ToCustomStartCell_StartsAtCorrectPosition()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("import_offset.xlsx", 1);
        var csvContent = "X,Y\n10,20";
        var inputPath = CreateTestCsvFile("import_offset.csv", csvContent);

        await service.ImportDataAsync(filePath, "Sheet1", inputPath, "C3");

        using var workbook = new XLWorkbook(filePath);
        var sheet = workbook.Worksheet("Sheet1");
        Assert.Equal("X", sheet.Cell("C3").GetValue<string>());
        Assert.Equal("Y", sheet.Cell("D3").GetValue<string>());
        Assert.Equal("10", sheet.Cell("C4").GetValue<string>());
    }

    [Fact]
    public async Task ImportDataAsync_WithUnsupportedFormat_ThrowsArgumentException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("import_unsupported.xlsx", 1);
        var xmlPath = CreateTestTextFile("data.xml", "<root></root>");

        await Assert.ThrowsAsync<ArgumentException>(() => service.ImportDataAsync(filePath, "Sheet1", xmlPath, "A1"));
    }

    [Fact]
    public async Task ImportDataAsync_CsvWithQuotes_ParsesCorrectly()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("import_quotes.xlsx", 1);
        var csvContent = "Name,Description\nItem,\"Hello, World\"";
        var inputPath = CreateTestCsvFile("import_quotes.csv", csvContent);

        await service.ImportDataAsync(filePath, "Sheet1", inputPath, "A1");

        using var workbook = new XLWorkbook(filePath);
        var sheet = workbook.Worksheet("Sheet1");
        Assert.Equal("Hello, World", sheet.Cell("B2").GetValue<string>());
    }

    [Fact]
    public async Task ImportDataAsync_EmptyJson_DoesNothing()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("import_empty_json.xlsx", 1);
        var jsonContent = "[]";
        var inputPath = CreateTestJsonFile("import_empty.json", jsonContent);

        await service.ImportDataAsync(filePath, "Sheet1", inputPath, "A1");

        // Should complete without error
        using var workbook = new XLWorkbook(filePath);
        Assert.NotNull(workbook.Worksheet("Sheet1"));
    }
}
