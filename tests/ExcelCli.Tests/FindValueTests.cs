using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for FindValueAsync operation
/// </summary>
public class FindValueTests : ExcelTestBase
{
    [Fact]
    public async Task FindValueAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.FindValueAsync(null!, "Sheet1", "value"));
    }

    [Fact]
    public async Task FindValueAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.FindValueAsync(nonExistentPath, "Sheet1", "value"));
    }

    [Fact]
    public async Task FindValueAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("find_nonexistent.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.FindValueAsync(filePath, "NonExistent", "value"));
    }

    [Fact]
    public async Task FindValueAsync_WhenValueExists_ReturnsLocation()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "Hello", "World" },
            new[] { "Test", "Data" }
        };
        var filePath = CreateTestExcelFileWithData("find_value.xlsx", "Sheet1", data);

        var results = (await service.FindValueAsync(filePath, "Sheet1", "World")).ToList();

        Assert.Single(results);
        Assert.Equal("Sheet1", results[0].SheetName);
        Assert.Equal("B1", results[0].CellAddress);
        Assert.Equal("World", results[0].Value);
    }

    [Fact]
    public async Task FindValueAsync_WhenValueNotExists_ReturnsEmptyList()
    {
        var service = CreateService();
        var data = new[] { new[] { "Hello", "World" } };
        var filePath = CreateTestExcelFileWithData("find_none.xlsx", "Sheet1", data);

        var results = (await service.FindValueAsync(filePath, "Sheet1", "NotHere")).ToList();

        Assert.Empty(results);
    }

    [Fact]
    public async Task FindValueAsync_WithFindAllFalse_ReturnsOnlyFirstMatch()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "Match", "Other" },
            new[] { "Match", "Other" }
        };
        var filePath = CreateTestExcelFileWithData("find_first.xlsx", "Sheet1", data);

        var results = (await service.FindValueAsync(filePath, "Sheet1", "Match", false)).ToList();

        Assert.Single(results);
        Assert.Equal("A1", results[0].CellAddress);
    }

    [Fact]
    public async Task FindValueAsync_WithFindAllTrue_ReturnsAllMatches()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "Match", "Other" },
            new[] { "Match", "Match" }
        };
        var filePath = CreateTestExcelFileWithData("find_all.xlsx", "Sheet1", data);

        var results = (await service.FindValueAsync(filePath, "Sheet1", "Match", true)).ToList();

        Assert.Equal(3, results.Count);
    }

    [Fact]
    public async Task FindValueAsync_IsCaseInsensitive()
    {
        var service = CreateService();
        var data = new[] { new[] { "HELLO", "hello", "Hello" } };
        var filePath = CreateTestExcelFileWithData("find_case.xlsx", "Sheet1", data);

        var results = (await service.FindValueAsync(filePath, "Sheet1", "HELLO", true)).ToList();

        Assert.Equal(3, results.Count);
    }

    [Fact]
    public async Task FindValueAsync_WithPartialMatch_FindsContaining()
    {
        var service = CreateService();
        var data = new[] { new[] { "HelloWorld", "Hello", "World" } };
        var filePath = CreateTestExcelFileWithData("find_partial.xlsx", "Sheet1", data);

        var results = (await service.FindValueAsync(filePath, "Sheet1", "Hello", true)).ToList();

        Assert.Equal(2, results.Count); // "HelloWorld" and "Hello"
    }

    [Fact]
    public async Task FindValueAsync_WithEmptySheet_ReturnsEmptyList()
    {
        var service = CreateService();
        var data = Array.Empty<string[]>();
        var filePath = CreateTestExcelFileWithData("find_truly_empty.xlsx", "Sheet1", data);

        var results = (await service.FindValueAsync(filePath, "Sheet1", "anything")).ToList();

        Assert.Empty(results);
    }
}
