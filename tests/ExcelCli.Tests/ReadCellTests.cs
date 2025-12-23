using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for ReadCellAsync operation
/// </summary>
public class ReadCellTests : ExcelTestBase
{
    [Fact]
    public async Task ReadCellAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadCellAsync(null!, "Sheet1", "A1"));
    }

    [Fact]
    public async Task ReadCellAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ReadCellAsync(nonExistentPath, "Sheet1", "A1"));
    }

    [Fact]
    public async Task ReadCellAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("read_cell.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.ReadCellAsync(filePath, "NonExistent", "A1"));
    }

    [Fact]
    public async Task ReadCellAsync_WithValidCell_ReturnsValue()
    {
        var service = CreateService();
        var data = new[] { new[] { "Hello", "World" } };
        var filePath = CreateTestExcelFileWithData("read_cell_data.xlsx", "Sheet1", data);

        var result = await service.ReadCellAsync(filePath, "Sheet1", "A1");

        Assert.Equal("Hello", result);
    }

    [Fact]
    public async Task ReadCellAsync_WithDifferentCell_ReturnsCorrectValue()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "A1", "B1", "C1" },
            new[] { "A2", "B2", "C2" }
        };
        var filePath = CreateTestExcelFileWithData("read_multiple_cells.xlsx", "Sheet1", data);

        Assert.Equal("A1", await service.ReadCellAsync(filePath, "Sheet1", "A1"));
        Assert.Equal("B1", await service.ReadCellAsync(filePath, "Sheet1", "B1"));
        Assert.Equal("C2", await service.ReadCellAsync(filePath, "Sheet1", "C2"));
    }

    [Fact]
    public async Task ReadCellAsync_WithEmptyCell_ReturnsEmptyString()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("empty_cell.xlsx", 1);

        var result = await service.ReadCellAsync(filePath, "Sheet1", "Z99");

        Assert.Equal("", result);
    }

    [Fact]
    public async Task ReadCellAsync_WithNumericValue_ReturnsStringRepresentation()
    {
        var service = CreateService();
        var data = new[] { new[] { "123", "456.78" } };
        var filePath = CreateTestExcelFileWithData("numeric_cell.xlsx", "Sheet1", data);

        var result = await service.ReadCellAsync(filePath, "Sheet1", "A1");

        Assert.Equal("123", result);
    }
}
