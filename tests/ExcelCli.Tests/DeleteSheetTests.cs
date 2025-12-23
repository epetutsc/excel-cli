using ClosedXML.Excel;
using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for DeleteSheetAsync operation
/// </summary>
public class DeleteSheetTests : ExcelTestBase
{
    [Fact]
    public async Task DeleteSheetAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.DeleteSheetAsync(null!, "Sheet1"));
    }

    [Fact]
    public async Task DeleteSheetAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.DeleteSheetAsync(nonExistentPath, "Sheet1"));
    }

    [Fact]
    public async Task DeleteSheetAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("delete_nonexistent.xlsx", 2);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.DeleteSheetAsync(filePath, "NonExistent"));
    }

    [Fact]
    public async Task DeleteSheetAsync_WithValidSheet_DeletesSheet()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("delete_sheet.xlsx", 2);

        await service.DeleteSheetAsync(filePath, "Sheet1");

        using var workbook = new XLWorkbook(filePath);
        Assert.False(workbook.Worksheets.Contains("Sheet1"));
        Assert.Single(workbook.Worksheets);
    }

    [Fact]
    public async Task DeleteSheetAsync_WithLastSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("single_sheet.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.DeleteSheetAsync(filePath, "Sheet1"));
    }

    [Fact]
    public async Task DeleteSheetAsync_DeletesCorrectSheet()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("delete_specific.xlsx", 3);

        await service.DeleteSheetAsync(filePath, "Sheet2");

        using var workbook = new XLWorkbook(filePath);
        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.True(workbook.Worksheets.Contains("Sheet1"));
        Assert.False(workbook.Worksheets.Contains("Sheet2"));
        Assert.True(workbook.Worksheets.Contains("Sheet3"));
    }

    [Fact]
    public async Task DeleteSheetAsync_CanDeleteMultipleSheets()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("multi_delete.xlsx", 4);

        await service.DeleteSheetAsync(filePath, "Sheet1");
        RefreshMockFile(filePath);
        await service.DeleteSheetAsync(filePath, "Sheet3");

        using var workbook = new XLWorkbook(filePath);
        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.True(workbook.Worksheets.Contains("Sheet2"));
        Assert.True(workbook.Worksheets.Contains("Sheet4"));
    }
}
