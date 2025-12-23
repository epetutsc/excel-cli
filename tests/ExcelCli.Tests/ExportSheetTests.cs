using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for ExportSheetAsync operation
/// </summary>
public class ExportSheetTests : ExcelTestBase
{
    [Fact]
    public async Task ExportSheetAsync_WithNullPath_ThrowsArgumentException()
    {
        var service = CreateService();
        var outputPath = Path.Combine(TestDirectory, "output.csv");

        await Assert.ThrowsAsync<ArgumentException>(() => service.ExportSheetAsync(null!, "Sheet1", outputPath, "csv"));
    }

    [Fact]
    public async Task ExportSheetAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var service = CreateService();
        var nonExistentPath = "/tmp/non-existent-file.xlsx";
        var outputPath = Path.Combine(TestDirectory, "output.csv");

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ExportSheetAsync(nonExistentPath, "Sheet1", outputPath, "csv"));
    }

    [Fact]
    public async Task ExportSheetAsync_WithNonExistentSheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var filePath = CreateTestExcelFile("export_nonexistent.xlsx", 1);
        var outputPath = Path.Combine(TestDirectory, "output.csv");

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.ExportSheetAsync(filePath, "NonExistent", outputPath, "csv"));
    }

    [Fact]
    public async Task ExportSheetAsync_WithEmptySheet_ThrowsInvalidOperationException()
    {
        var service = CreateService();
        var data = Array.Empty<string[]>();
        var filePath = CreateTestExcelFileWithData("export_empty.xlsx", "Sheet1", data);
        var outputPath = Path.Combine(TestDirectory, "export_empty.csv");

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.ExportSheetAsync(filePath, "Sheet1", outputPath, "csv"));
    }

    [Fact]
    public async Task ExportSheetAsync_ToCsv_CreatesFile()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "Name", "Age" },
            new[] { "Alice", "30" }
        };
        var filePath = CreateTestExcelFileWithData("export_csv.xlsx", "Sheet1", data);
        var outputPath = Path.Combine(TestDirectory, "export_csv.csv");

        await service.ExportSheetAsync(filePath, "Sheet1", outputPath, "csv");

        Assert.True(FileSystem.File.Exists(outputPath));
        var content = await FileSystem.File.ReadAllTextAsync(outputPath);
        Assert.Contains("Name,Age", content);
        Assert.Contains("Alice,30", content);
    }

    [Fact]
    public async Task ExportSheetAsync_ToJson_CreatesFile()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "Name", "Age" },
            new[] { "Bob", "25" }
        };
        var filePath = CreateTestExcelFileWithData("export_json.xlsx", "Sheet1", data);
        var outputPath = Path.Combine(TestDirectory, "export_json.json");

        await service.ExportSheetAsync(filePath, "Sheet1", outputPath, "json");

        Assert.True(FileSystem.File.Exists(outputPath));
        var content = await FileSystem.File.ReadAllTextAsync(outputPath);
        Assert.Contains("\"Name\"", content);
        Assert.Contains("\"Bob\"", content);
        Assert.Contains("\"Age\"", content);
        Assert.Contains("\"25\"", content);
    }

    [Fact]
    public async Task ExportSheetAsync_WithUnsupportedFormat_ThrowsArgumentException()
    {
        var service = CreateService();
        var data = new[] { new[] { "Data" } };
        var filePath = CreateTestExcelFileWithData("export_unsupported.xlsx", "Sheet1", data);
        var outputPath = Path.Combine(TestDirectory, "export.xml");

        await Assert.ThrowsAsync<ArgumentException>(() => service.ExportSheetAsync(filePath, "Sheet1", outputPath, "xml"));
    }

    [Fact]
    public async Task ExportSheetAsync_ToCsvWithComma_EscapesValue()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "Name", "Description" },
            new[] { "Item", "Hello, World" }
        };
        var filePath = CreateTestExcelFileWithData("export_escape.xlsx", "Sheet1", data);
        var outputPath = Path.Combine(TestDirectory, "export_escape.csv");

        await service.ExportSheetAsync(filePath, "Sheet1", outputPath, "csv");

        var content = await FileSystem.File.ReadAllTextAsync(outputPath);
        Assert.Contains("\"Hello, World\"", content);
    }

    [Fact]
    public async Task ExportSheetAsync_JsonWithMultipleRows_CreatesArray()
    {
        var service = CreateService();
        var data = new[]
        {
            new[] { "Id", "Value" },
            new[] { "1", "A" },
            new[] { "2", "B" },
            new[] { "3", "C" }
        };
        var filePath = CreateTestExcelFileWithData("export_multi_json.xlsx", "Sheet1", data);
        var outputPath = Path.Combine(TestDirectory, "export_multi.json");

        await service.ExportSheetAsync(filePath, "Sheet1", outputPath, "json");

        var content = await FileSystem.File.ReadAllTextAsync(outputPath);
        Assert.Contains("[", content); // Array start
        Assert.Contains("]", content); // Array end
        // Should have 3 data rows (excluding header)
        Assert.Contains("\"1\"", content);
        Assert.Contains("\"2\"", content);
        Assert.Contains("\"3\"", content);
    }
}
