using ClosedXML.Excel;
using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for WriteCellAsync operation
/// </summary>
public class WriteCellTests : ExcelTestBase
{
    [Fact]
    public async Task WriteCellAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.WriteCellAsync(null!, "Sheet1", "A1", "value"));
    }

    [Fact]
    public async Task WriteCellAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.WriteCellAsync(nonExistentPath, "Sheet1", "A1", "value"));
    }

    [Fact]
    public async Task WriteCellAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("write_cell.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.WriteCellAsync(filePath, "NonExistent", "A1", "value"));
    }

    [Fact]
    public async Task WriteCellAsync_WithValidCell_WritesValue()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("write_valid.xlsx", 1);

        await service.WriteCellAsync(filePath, "Sheet1", "A1", "NewValue");

        // Verify by reading the file directly
        using var workbook = new XLWorkbook(filePath);
        var value = workbook.Worksheet("Sheet1").Cell("A1").GetValue<string>();
        Assert.Equal("NewValue", value);
    }

    [Fact]
    public async Task WriteCellAsync_OverwritesExistingValue()
    {
        var service = CreateService();
        var data = new[] { new[] { "OldValue" } };
        var filePath = CreateTestExcelFileWithData("overwrite.xlsx", "Sheet1", data);

        await service.WriteCellAsync(filePath, "Sheet1", "A1", "NewValue");

        using var workbook = new XLWorkbook(filePath);
        var value = workbook.Worksheet("Sheet1").Cell("A1").GetValue<string>();
        Assert.Equal("NewValue", value);
    }

    [Fact]
    public async Task WriteCellAsync_WritesToDifferentCells()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("multi_write.xlsx", 1);

        await service.WriteCellAsync(filePath, "Sheet1", "A1", "Value1");
        RefreshMockFile(filePath);
        await service.WriteCellAsync(filePath, "Sheet1", "B2", "Value2");
        RefreshMockFile(filePath);
        await service.WriteCellAsync(filePath, "Sheet1", "C3", "Value3");

        using var workbook = new XLWorkbook(filePath);
        var sheet = workbook.Worksheet("Sheet1");
        Assert.Equal("Value1", sheet.Cell("A1").GetValue<string>());
        Assert.Equal("Value2", sheet.Cell("B2").GetValue<string>());
        Assert.Equal("Value3", sheet.Cell("C3").GetValue<string>());
    }

    [Fact]
    public async Task WriteCellAsync_WithEmptyValue_WritesEmptyString()
    {
        var service = CreateService();
        var data = new[] { new[] { "HasValue" } };
        var filePath = CreateTestExcelFileWithData("empty_write.xlsx", "Sheet1", data);

        await service.WriteCellAsync(filePath, "Sheet1", "A1", "");

        using var workbook = new XLWorkbook(filePath);
        var value = workbook.Worksheet("Sheet1").Cell("A1").GetValue<string>();
        Assert.Equal("", value);
    }

    [Fact]
    public async Task WriteCellAsync_WithSpecialCharacters_WritesCorrectly()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("special_chars.xlsx", 1);

        await service.WriteCellAsync(filePath, "Sheet1", "A1", "Hello, \"World\"!");

        using var workbook = new XLWorkbook(filePath);
        var value = workbook.Worksheet("Sheet1").Cell("A1").GetValue<string>();
        Assert.Equal("Hello, \"World\"!", value);
    }
}
