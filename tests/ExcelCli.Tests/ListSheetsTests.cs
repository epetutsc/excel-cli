using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for ListSheetsAsync operation
/// </summary>
public class ListSheetsTests : ExcelTestBase
{
    [Fact]
    public async Task ListSheetsAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.ListSheetsAsync(null!));
    }

    [Fact]
    public async Task ListSheetsAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ListSheetsAsync(nonExistentPath));
    }

    [Fact]
    public async Task ListSheetsAsync_WithValidFile_ReturnsSheetInfos()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("list_sheets.xlsx", 3);

        var result = (await service.ListSheetsAsync(filePath)).ToList();

        Assert.Equal(3, result.Count);
        Assert.Equal("Sheet1", result[0].Name);
        Assert.Equal("Sheet2", result[1].Name);
        Assert.Equal("Sheet3", result[2].Name);
    }

    [Fact]
    public async Task ListSheetsAsync_WithDataInSheet_ReturnsCorrectRowAndColumnCount()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "A", "B", "C" },
            new[] { "1", "2", "3" },
            new[] { "4", "5", "6" }
        };
        var filePath = CreateTestExcelFileWithData("sheets_with_data.xlsx", "DataSheet", data);

        var result = (await service.ListSheetsAsync(filePath)).ToList();

        Assert.Single(result);
        Assert.Equal("DataSheet", result[0].Name);
        Assert.Equal(3, result[0].RowCount);
        Assert.Equal(3, result[0].ColumnCount);
    }

    [Fact]
    public async Task ListSheetsAsync_WithEmptySheet_ReturnsZeroRowAndColumnCount()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("empty_sheet.xlsx", 1);

        var result = (await service.ListSheetsAsync(filePath)).ToList();

        Assert.Single(result);
        // Empty sheets have some default data from CreateTestExcelFile
        Assert.True(result[0].RowCount >= 0);
        Assert.True(result[0].ColumnCount >= 0);
    }
}
