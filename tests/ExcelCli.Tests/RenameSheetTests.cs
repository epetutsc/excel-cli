using ClosedXML.Excel;
using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for RenameSheetAsync operation
/// </summary>
public class RenameSheetTests : ExcelTestBase
{
    [Fact]
    public async Task RenameSheetAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.RenameSheetAsync(null!, "OldName", "NewName"));
    }

    [Fact]
    public async Task RenameSheetAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.RenameSheetAsync(nonExistentPath, "Sheet1", "NewName"));
    }

    [Fact]
    public async Task RenameSheetAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("rename_nonexistent.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.RenameSheetAsync(filePath, "NonExistent", "NewName"));
    }

    [Fact]
    public async Task RenameSheetAsync_WithValidNames_RenamesSheet()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("rename_sheet.xlsx", 1);

        await service.RenameSheetAsync(filePath, "Sheet1", "RenamedSheet");

        using var workbook = new XLWorkbook(filePath);
        Assert.False(workbook.Worksheets.Contains("Sheet1"));
        Assert.True(workbook.Worksheets.Contains("RenamedSheet"));
    }

    [Fact]
    public async Task RenameSheetAsync_ToExistingName_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("rename_duplicate.xlsx", 2);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.RenameSheetAsync(filePath, "Sheet1", "Sheet2"));
    }

    [Fact]
    public async Task RenameSheetAsync_PreservesData()
    {
        var service = CreateService();
        var data = new[] { new[] { "TestData" } };
        var filePath = CreateTestExcelFileWithData("rename_preserve.xlsx", "Sheet1", data);

        await service.RenameSheetAsync(filePath, "Sheet1", "NewName");

        using var workbook = new XLWorkbook(filePath);
        var value = workbook.Worksheet("NewName").Cell("A1").GetValue<string>();
        Assert.Equal("TestData", value);
    }

    [Fact]
    public async Task RenameSheetAsync_WithSpecialCharacters_Works()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("rename_special.xlsx", 1);

        await service.RenameSheetAsync(filePath, "Sheet1", "Data 2024-Q1");

        using var workbook = new XLWorkbook(filePath);
        Assert.True(workbook.Worksheets.Contains("Data 2024-Q1"));
    }
}
