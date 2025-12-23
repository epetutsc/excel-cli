using ExcelCli.Services;
using NSubstitute;
using Serilog;
using System.IO.Abstractions.TestingHelpers;
using Xunit;

namespace ExcelCli.Tests;

public class ExcelServiceTests
{
    private readonly ILogger _logger;
    private readonly MockFileSystem _fileSystem;

    public ExcelServiceTests()
    {
        _logger = Substitute.For<ILogger>();
        _fileSystem = new MockFileSystem();
    }

    [Fact]
    public void Constructor_WithNullLogger_ThrowsArgumentNullException()
    {
        // Arrange & Act & Assert
        Assert.Throws<ArgumentNullException>(() => new ExcelService(null!, _fileSystem));
    }

    [Fact]
    public void Constructor_WithNullFileSystem_ThrowsArgumentNullException()
    {
        // Arrange & Act & Assert
        Assert.Throws<ArgumentNullException>(() => new ExcelService(_logger, null!));
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadFileInfoAsync(null!));
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithEmptyPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadFileInfoAsync(""));
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ReadFileInfoAsync(nonExistentPath));
    }

    [Fact]
    public async Task ListSheetsAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.ListSheetsAsync(null!));
    }

    [Fact]
    public async Task ListSheetsAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ListSheetsAsync(nonExistentPath));
    }

    [Fact]
    public async Task ReadCellAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadCellAsync(null!, "Sheet1", "A1"));
    }

    [Fact]
    public async Task ReadCellAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ReadCellAsync(nonExistentPath, "Sheet1", "A1"));
    }

    [Fact]
    public async Task ReadRangeAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadRangeAsync(null!, "Sheet1", "A1:B2"));
    }

    [Fact]
    public async Task WriteCellAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.WriteCellAsync(null!, "Sheet1", "A1", "value"));
    }

    [Fact]
    public async Task CreateSheetAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.CreateSheetAsync(null!, "NewSheet"));
    }

    [Fact]
    public async Task DeleteSheetAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.DeleteSheetAsync(null!, "Sheet1"));
    }

    [Fact]
    public async Task RenameSheetAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.RenameSheetAsync(null!, "OldName", "NewName"));
    }

    [Fact]
    public async Task CopySheetAsync_WithNullSourcePath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.CopySheetAsync(null!, "Sheet1", "/tmp/target.xlsx"));
    }

    [Fact]
    public async Task FindValueAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.FindValueAsync(null!, "Sheet1", "value"));
    }

    [Fact]
    public async Task ExportSheetAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.ExportSheetAsync(null!, "Sheet1", "/tmp/output.csv", "csv"));
    }

    [Fact]
    public async Task ImportDataAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.ImportDataAsync(null!, "Sheet1", "/tmp/input.csv", "A1"));
    }

    [Fact]
    public async Task InsertFormulaAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger, _fileSystem);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.InsertFormulaAsync(null!, "Sheet1", "A1", "=SUM(A1:B1)"));
    }
}
