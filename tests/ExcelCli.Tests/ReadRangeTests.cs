using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for ReadRangeAsync operation
/// </summary>
public class ReadRangeTests : ExcelTestBase
{
    [Fact]
    public async Task ReadRangeAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadRangeAsync(null!, "Sheet1", "A1:B2"));
    }

    [Fact]
    public async Task ReadRangeAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ReadRangeAsync(nonExistentPath, "Sheet1", "A1:B2"));
    }

    [Fact]
    public async Task ReadRangeAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("read_range.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.ReadRangeAsync(filePath, "NonExistent", "A1:B2"));
    }

    [Fact]
    public async Task ReadRangeAsync_WithValidRange_ReturnsData()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "A1", "B1" },
            new[] { "A2", "B2" }
        };
        var filePath = CreateTestExcelFileWithData("range_data.xlsx", "Sheet1", data);

        var result = await service.ReadRangeAsync(filePath, "Sheet1", "A1:B2");

        Assert.Equal(2, result.Length);
        Assert.Equal(2, result[0].Length);
        Assert.Equal("A1", result[0][0]);
        Assert.Equal("B1", result[0][1]);
        Assert.Equal("A2", result[1][0]);
        Assert.Equal("B2", result[1][1]);
    }

    [Fact]
    public async Task ReadRangeAsync_WithLargerRange_ReturnsAllData()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "1", "2", "3", "4" },
            new[] { "5", "6", "7", "8" },
            new[] { "9", "10", "11", "12" }
        };
        var filePath = CreateTestExcelFileWithData("larger_range.xlsx", "Sheet1", data);

        var result = await service.ReadRangeAsync(filePath, "Sheet1", "A1:D3");

        Assert.Equal(3, result.Length);
        Assert.Equal(4, result[0].Length);
        Assert.Equal("1", result[0][0]);
        Assert.Equal("12", result[2][3]);
    }

    [Fact]
    public async Task ReadRangeAsync_WithSingleCellRange_ReturnsOneElement()
    {
        var service = CreateService();
        var data = new[] { new[] { "Single" } };
        var filePath = CreateTestExcelFileWithData("single_range.xlsx", "Sheet1", data);

        var result = await service.ReadRangeAsync(filePath, "Sheet1", "A1:A1");

        Assert.Single(result);
        Assert.Single(result[0]);
        Assert.Equal("Single", result[0][0]);
    }

    [Fact]
    public async Task ReadRangeAsync_WithPartiallyFilledRange_ReturnsEmptyForUnfilled()
    {
        var service = CreateService();
        var data = new[] { new[] { "A1" } };
        var filePath = CreateTestExcelFileWithData("partial_range.xlsx", "Sheet1", data);

        var result = await service.ReadRangeAsync(filePath, "Sheet1", "A1:B2");

        Assert.Equal(2, result.Length);
        Assert.Equal(2, result[0].Length);
        Assert.Equal("A1", result[0][0]);
        Assert.Equal("", result[0][1]);
        Assert.Equal("", result[1][0]);
        Assert.Equal("", result[1][1]);
    }
}
