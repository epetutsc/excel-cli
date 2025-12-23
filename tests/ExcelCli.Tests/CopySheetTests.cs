using ClosedXML.Excel;
using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for CopySheetAsync operation
/// </summary>
public class CopySheetTests : ExcelTestBase
{
    [Fact]
    public async Task CopySheetAsync_WithNullSourcePath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.CopySheetAsync(null!, "Sheet1", "/tmp/target.xlsx"));
    }

    [Fact]
    public async Task CopySheetAsync_WithNonExistentSourceFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";
        var targetPath = Path.Combine(TestDirectory, "target.xlsx");

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.CopySheetAsync(nonExistentPath, "Sheet1", targetPath));
    }

    [Fact]
    public async Task CopySheetAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var sourcePath = CreateTestExcelFile("copy_source.xlsx", 1);
        var targetPath = Path.Combine(TestDirectory, "copy_target.xlsx");

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.CopySheetAsync(sourcePath, "NonExistent", targetPath));
    }

    [Fact]
    public async Task CopySheetAsync_ToNewFile_CreatesFileAndCopiesSheet()
    {
        var service = CreateService();
        var sourcePath = CreateTestExcelFile("copy_source_new.xlsx", 1);
        var targetPath = Path.Combine(TestDirectory, "copy_new_target.xlsx");

        await service.CopySheetAsync(sourcePath, "Sheet1", targetPath);

        Assert.True(File.Exists(targetPath));
        using var targetWorkbook = new XLWorkbook(targetPath);
        Assert.True(targetWorkbook.Worksheets.Contains("Sheet1"));
    }

    [Fact]
    public async Task CopySheetAsync_ToExistingFile_AddsSheet()
    {
        var service = CreateService();
        var sourcePath = CreateTestExcelFile("copy_source_existing.xlsx", 1);
        var targetPath = CreateTestExcelFile("copy_existing_target.xlsx", 1);

        await service.CopySheetAsync(sourcePath, "Sheet1", targetPath, "CopiedSheet");

        using var targetWorkbook = new XLWorkbook(targetPath);
        Assert.Equal(2, targetWorkbook.Worksheets.Count);
        Assert.True(targetWorkbook.Worksheets.Contains("Sheet1"));
        Assert.True(targetWorkbook.Worksheets.Contains("CopiedSheet"));
    }

    [Fact]
    public async Task CopySheetAsync_WithNewName_RenamesCopiedSheet()
    {
        var service = CreateService();
        var sourcePath = CreateTestExcelFile("copy_rename.xlsx", 1);
        var targetPath = Path.Combine(TestDirectory, "copy_rename_target.xlsx");

        await service.CopySheetAsync(sourcePath, "Sheet1", targetPath, "CustomName");

        using var targetWorkbook = new XLWorkbook(targetPath);
        Assert.True(targetWorkbook.Worksheets.Contains("CustomName"));
        Assert.False(targetWorkbook.Worksheets.Contains("Sheet1"));
    }

    [Fact]
    public async Task CopySheetAsync_PreservesData()
    {
        var service = CreateService();
        var data = new[] { new[] { "CopiedData", "123" } };
        var sourcePath = CreateTestExcelFileWithData("copy_data_source.xlsx", "DataSheet", data);
        var targetPath = Path.Combine(TestDirectory, "copy_data_target.xlsx");

        await service.CopySheetAsync(sourcePath, "DataSheet", targetPath);

        using var targetWorkbook = new XLWorkbook(targetPath);
        var value = targetWorkbook.Worksheet("DataSheet").Cell("A1").GetValue<string>();
        Assert.Equal("CopiedData", value);
    }

    [Fact]
    public async Task CopySheetAsync_WithoutNewName_KeepsOriginalName()
    {
        var service = CreateService();
        var sourcePath = CreateTestExcelFile("copy_original_name.xlsx", 1);
        var targetPath = Path.Combine(TestDirectory, "copy_original_target.xlsx");

        await service.CopySheetAsync(sourcePath, "Sheet1", targetPath, null);

        using var targetWorkbook = new XLWorkbook(targetPath);
        Assert.True(targetWorkbook.Worksheets.Contains("Sheet1"));
    }
}
