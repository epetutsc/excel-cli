using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for ReadFileInfoAsync operation
/// </summary>
public class ReadFileInfoTests : ExcelTestBase
{
    [Fact]
    public async Task ReadFileInfoAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadFileInfoAsync(null!));
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithEmptyPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadFileInfoAsync(""));
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ReadFileInfoAsync(nonExistentPath));
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithValidFile_ReturnsFileInfo()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("test_read_info.xlsx", 3);

        var result = await service.ReadFileInfoAsync(filePath);

        Assert.Equal("test_read_info.xlsx", result.FileName);
        Assert.Equal(3, result.SheetCount);
        Assert.True(result.FileSize > 0);
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithSingleSheetFile_ReturnsCorrectSheetCount()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("single_sheet.xlsx", 1);

        var result = await service.ReadFileInfoAsync(filePath);

        Assert.Equal(1, result.SheetCount);
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithMultipleSheetFile_ReturnsCorrectSheetCount()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("multi_sheet.xlsx", 5);

        var result = await service.ReadFileInfoAsync(filePath);

        Assert.Equal(5, result.SheetCount);
    }
}
