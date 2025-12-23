using ClosedXML.Excel;
using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for CreateSheetAsync operation
/// </summary>
public class CreateSheetTests : ExcelTestBase
{
    [Fact]
    public async Task CreateSheetAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.CreateSheetAsync(null!, "NewSheet"));
    }

    [Fact]
    public async Task CreateSheetAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.CreateSheetAsync(nonExistentPath, "NewSheet"));
    }

    [Fact]
    public async Task CreateSheetAsync_WithValidName_CreatesSheet()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("create_sheet.xlsx", 1);

        await service.CreateSheetAsync(filePath, "NewSheet");

        using var workbook = new XLWorkbook(filePath);
        Assert.True(workbook.Worksheets.Contains("NewSheet"));
        Assert.Equal(2, workbook.Worksheets.Count);
    }

    [Fact]
    public async Task CreateSheetAsync_WithExistingName_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("duplicate_sheet.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.CreateSheetAsync(filePath, "Sheet1"));
    }

    [Fact]
    public async Task CreateSheetAsync_CreatesMultipleSheets()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("multi_create.xlsx", 1);

        await service.CreateSheetAsync(filePath, "Sheet2");
        RefreshMockFile(filePath);
        await service.CreateSheetAsync(filePath, "Sheet3");

        using var workbook = new XLWorkbook(filePath);
        Assert.Equal(3, workbook.Worksheets.Count);
        Assert.True(workbook.Worksheets.Contains("Sheet1"));
        Assert.True(workbook.Worksheets.Contains("Sheet2"));
        Assert.True(workbook.Worksheets.Contains("Sheet3"));
    }

    [Fact]
    public async Task CreateSheetAsync_WithSpecialCharactersInName_CreatesSheet()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("special_name.xlsx", 1);

        await service.CreateSheetAsync(filePath, "Data 2024");

        using var workbook = new XLWorkbook(filePath);
        Assert.True(workbook.Worksheets.Contains("Data 2024"));
    }
}
