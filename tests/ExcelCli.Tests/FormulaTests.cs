using ClosedXML.Excel;
using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for formula operations (read, calculate, set)
/// This tests InsertFormulaAsync and related formula functionality
/// </summary>
public class FormulaTests : ExcelTestBase
{
    [Fact]
    public async Task InsertFormulaAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();

        await Assert.ThrowsAsync<ArgumentException>(() => service.InsertFormulaAsync(null!, "Sheet1", "A1", "=SUM(A1:B1)"));
    }

    [Fact]
    public async Task InsertFormulaAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.InsertFormulaAsync(nonExistentPath, "Sheet1", "A1", "=SUM(A1:B1)"));
    }

    [Fact]
    public async Task InsertFormulaAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("formula_no_sheet.xlsx", 1);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.InsertFormulaAsync(filePath, "NonExistent", "A1", "=SUM(A1:B1)"));
    }

    [Fact]
    public async Task InsertFormulaAsync_WithValidFormula_SetsFormula()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("formula_set.xlsx", 1);

        await service.InsertFormulaAsync(filePath, "Sheet1", "C1", "=A1+B1");

        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("C1");
        Assert.True(cell.HasFormula);
        Assert.Equal("A1+B1", cell.FormulaA1);
    }

    [Fact]
    public async Task InsertFormulaAsync_WithoutEqualsSign_AddsEqualsSign()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("formula_no_equals.xlsx", 1);

        await service.InsertFormulaAsync(filePath, "Sheet1", "C1", "A1+B1");

        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("C1");
        Assert.True(cell.HasFormula);
        Assert.Equal("A1+B1", cell.FormulaA1);
    }

    [Fact]
    public void Formula_Sum_CalculatesCorrectly()
    {
        var filePath = CreateTestExcelFileWithFormulas("formula_sum.xlsx", "Sheet1");

        // The file has A1=10, A2=20, A3=30, and C3 already has =SUM(A1:A3)
        // Let's verify by reading the formula
        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("C3");
        Assert.True(cell.HasFormula);
        Assert.Equal("SUM(A1:A3)", cell.FormulaA1);
        
        // ClosedXML calculates the value
        var value = cell.Value;
        Assert.Equal(60.0, value.GetNumber());
    }

    [Fact]
    public void Formula_Multiplication_CalculatesCorrectly()
    {
        var filePath = CreateTestExcelFileWithFormulas("formula_multiply.xlsx", "Sheet1");

        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("C2");
        Assert.True(cell.HasFormula);
        Assert.Equal("A2*B2", cell.FormulaA1);
        
        // A2=20, B2=15, so result should be 300
        var value = cell.Value;
        Assert.Equal(300.0, value.GetNumber());
    }

    [Fact]
    public void Formula_Average_CalculatesCorrectly()
    {
        var filePath = CreateTestExcelFileWithFormulas("formula_average.xlsx", "Sheet1");

        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("D1");
        Assert.True(cell.HasFormula);
        Assert.Equal("AVERAGE(B1:B3)", cell.FormulaA1);
        
        // B1=5, B2=15, B3=25, average = 15
        var value = cell.Value;
        Assert.Equal(15.0, value.GetNumber());
    }

    [Fact]
    public void ReadFormula_FromCell_ReturnsFormulaString()
    {
        var filePath = CreateTestExcelFileWithFormulas("read_formula.xlsx", "Sheet1");

        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("C1");
        
        // Verify we can read the formula
        Assert.True(cell.HasFormula);
        Assert.Equal("A1+B1", cell.FormulaA1);
    }

    [Fact]
    public void ReadCellWithFormula_GetCalculatedValue_ReturnsNumericResult()
    {
        var filePath = CreateTestExcelFileWithFormulas("read_calculated.xlsx", "Sheet1");

        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("C1");
        
        // C1 has formula =A1+B1 where A1=10 and B1=5
        Assert.True(cell.HasFormula);
        var value = cell.Value;
        Assert.Equal(15.0, value.GetNumber());
    }

    [Fact]
    public async Task InsertFormulaAsync_OverwritesExistingValue()
    {
        var service = CreateService();
        var data = new[] { new[] { "10", "20", "OldValue" } };
        var filePath = CreateTestExcelFileWithData("formula_overwrite.xlsx", "Sheet1", data);

        await service.InsertFormulaAsync(filePath, "Sheet1", "C1", "=A1+B1");

        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("C1");
        Assert.True(cell.HasFormula);
        // Value should now be calculated from formula
    }

    [Fact]
    public async Task InsertFormulaAsync_WithComplexFormula_Works()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("formula_complex.xlsx", 1);

        await service.InsertFormulaAsync(filePath, "Sheet1", "D1", "=IF(A1>10,\"High\",\"Low\")");

        using var workbook = new XLWorkbook(filePath);
        var cell = workbook.Worksheet("Sheet1").Cell("D1");
        Assert.True(cell.HasFormula);
        Assert.Equal("IF(A1>10,\"High\",\"Low\")", cell.FormulaA1);
    }

    [Fact]
    public async Task ReadCellAsync_WithFormula_ReturnsFormula()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFileWithFormulas("read_formula_value.xlsx", "Sheet1");
        RefreshMockFile(filePath);

        // Use the service to read the cell - should return formula, not calculated value
        var result = await service.ReadCellAsync(filePath, "Sheet1", "C1");

        // C1 has formula =A1+B1 where A1=10 and B1=5
        // ReadCellAsync should now return the formula "=A1+B1"
        Assert.Equal("=A1+B1", result);
    }

    [Fact]
    public async Task ReadRangeAsync_WithFormulas_ReturnsCalculatedValues()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFileWithFormulas("read_range_formula.xlsx", "Sheet1");
        RefreshMockFile(filePath);

        var result = await service.ReadRangeAsync(filePath, "Sheet1", "C1:C3");

        // C1 = A1+B1 = 10+5 = 15
        // C2 = A2*B2 = 20*15 = 300
        // C3 = SUM(A1:A3) = 10+20+30 = 60
        Assert.Equal(3, result.Length);
        Assert.Equal("15", result[0][0]);
        Assert.Equal("300", result[1][0]);
        Assert.Equal("60", result[2][0]);
    }

    [Fact]
    public void CellWithFormula_HasFormulaProperty_IsTrue()
    {
        var filePath = CreateTestExcelFileWithFormulas("has_formula.xlsx", "Sheet1");

        using var workbook = new XLWorkbook(filePath);
        var sheet = workbook.Worksheet("Sheet1");
        
        // Cells with formulas
        Assert.True(sheet.Cell("C1").HasFormula);
        Assert.True(sheet.Cell("C2").HasFormula);
        Assert.True(sheet.Cell("C3").HasFormula);
        Assert.True(sheet.Cell("D1").HasFormula);
        
        // Cells without formulas
        Assert.False(sheet.Cell("A1").HasFormula);
        Assert.False(sheet.Cell("B1").HasFormula);
    }

    [Fact]
    public void CellWithFormula_ChangingReferencedCell_RecalculatesValue()
    {
        var filePath = CreateTestExcelFileWithFormulas("recalculate.xlsx", "Sheet1");

        using var workbook = new XLWorkbook(filePath);
        var sheet = workbook.Worksheet("Sheet1");
        
        // Original: A1=10, C1=A1+B1=15
        Assert.Equal(15.0, sheet.Cell("C1").Value.GetNumber());
        
        // Change A1
        sheet.Cell("A1").Value = 100;
        
        // C1 should recalculate
        Assert.Equal(105.0, sheet.Cell("C1").Value.GetNumber());
    }
}
