using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for GetCellValueAsync operation
/// </summary>
public class ReadCellValueTests : ExcelTestBase
{
    [Fact]
    public async Task GetCellValueAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.GetCellValueAsync(null!, "Sheet1", "A1"));
    }

    [Fact]
    public async Task GetCellValueAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.GetCellValueAsync(nonExistentPath, "Sheet1", "A1"));
    }

    [Fact]
    public async Task GetCellValueAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("get_cell_value.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.GetCellValueAsync(filePath, "NonExistent", "A1"));
    }

    [Fact]
    public async Task GetCellValueAsync_WithValidCell_ReturnsValue()
    {
        var service = CreateService();
        var data = new[] { new[] { "Hello", "World" } };
        var filePath = CreateTestExcelFileWithData("get_cell_value_data.xlsx", "Sheet1", data);

        var result = await service.GetCellValueAsync(filePath, "Sheet1", "A1");

        Assert.Equal("Hello", result);
    }

    [Fact]
    public async Task GetCellValueAsync_WithDifferentCell_ReturnsCorrectValue()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "A1", "B1", "C1" },
            new[] { "A2", "B2", "C2" }
        };
        var filePath = CreateTestExcelFileWithData("get_multiple_cells.xlsx", "Sheet1", data);

        Assert.Equal("A1", await service.GetCellValueAsync(filePath, "Sheet1", "A1"));
        Assert.Equal("B1", await service.GetCellValueAsync(filePath, "Sheet1", "B1"));
        Assert.Equal("C2", await service.GetCellValueAsync(filePath, "Sheet1", "C2"));
    }

    [Fact]
    public async Task GetCellValueAsync_WithEmptyCell_ReturnsEmptyString()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("empty_cell_value.xlsx", 1);

        var result = await service.GetCellValueAsync(filePath, "Sheet1", "Z99");

        Assert.Equal("", result);
    }

    [Fact]
    public async Task GetCellValueAsync_WithNumericValue_ReturnsStringRepresentation()
    {
        var service = CreateService();
        var data = new[] { new[] { "123", "456.78" } };
        var filePath = CreateTestExcelFileWithData("numeric_cell_value.xlsx", "Sheet1", data);

        var result = await service.GetCellValueAsync(filePath, "Sheet1", "A1");

        Assert.Equal("123", result);
    }

    [Fact]
    public async Task GetCellValueAsync_WithFormulaCell_ReturnsCalculatedValue()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFileWithFormulas("get_formula_value.xlsx", "Sheet1");
        RefreshMockFile(filePath);

        // C1 has formula =A1+B1 where A1=10 and B1=5
        var result = await service.GetCellValueAsync(filePath, "Sheet1", "C1");

        // Should return the calculated value "15", not the formula "=A1+B1"
        Assert.Equal("15", result);
    }

    [Fact]
    public async Task GetCellValueAsync_WithSumFormula_ReturnsCalculatedSum()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFileWithFormulas("get_sum_value.xlsx", "Sheet1");
        RefreshMockFile(filePath);

        // C3 has formula =SUM(A1:A3) where A1=10, A2=20, A3=30
        var result = await service.GetCellValueAsync(filePath, "Sheet1", "C3");

        // Should return "60"
        Assert.Equal("60", result);
    }

    [Fact]
    public async Task GetCellValueAsync_WithMultiplicationFormula_ReturnsCalculatedProduct()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFileWithFormulas("get_multiply_value.xlsx", "Sheet1");
        RefreshMockFile(filePath);

        // C2 has formula =A2*B2 where A2=20 and B2=15
        var result = await service.GetCellValueAsync(filePath, "Sheet1", "C2");

        // Should return "300"
        Assert.Equal("300", result);
    }

    [Fact]
    public async Task GetCellValueAsync_WithAverageFormula_ReturnsCalculatedAverage()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFileWithFormulas("get_average_value.xlsx", "Sheet1");
        RefreshMockFile(filePath);

        // D1 has formula =AVERAGE(B1:B3) where B1=5, B2=15, B3=25
        var result = await service.GetCellValueAsync(filePath, "Sheet1", "D1");

        // Should return "15"
        Assert.Equal("15", result);
    }

    [Fact]
    public async Task GetCellValueAsync_WithComplexFormula_ReturnsCalculatedResult()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFileWithFormulas("get_complex_value.xlsx", "Sheet1");
        RefreshMockFile(filePath);

        // C1 has formula =A1+B1 where A1=10 and B1=5, result should be 15
        var c1Result = await service.GetCellValueAsync(filePath, "Sheet1", "C1");
        Assert.Equal("15", c1Result);

        // C2 has formula =A2*B2 where A2=20 and B2=15, result should be 300
        var c2Result = await service.GetCellValueAsync(filePath, "Sheet1", "C2");
        Assert.Equal("300", c2Result);

        // C3 has formula =SUM(A1:A3) where A1=10, A2=20, A3=30, result should be 60
        var c3Result = await service.GetCellValueAsync(filePath, "Sheet1", "C3");
        Assert.Equal("60", c3Result);
    }
}
