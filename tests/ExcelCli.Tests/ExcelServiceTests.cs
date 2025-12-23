using ExcelCli.Services;
using NSubstitute;
using Serilog;
using Xunit;

namespace ExcelCli.Tests;

public class ExcelServiceTests
{
    private readonly ILogger _logger;

    public ExcelServiceTests()
    {
        _logger = Substitute.For<ILogger>();
    }

    [Fact]
    public void Constructor_WithNullLogger_ThrowsArgumentNullException()
    {
        // Arrange & Act & Assert
        Assert.Throws<ArgumentNullException>(() => new ExcelService(null!));
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithNullPath_ThrowsArgumentException()
    {
        // Arrange
        var service = new ExcelService(_logger);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => service.ReadFileInfoAsync(null!));
    }

    [Fact]
    public async Task ReadFileInfoAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        var service = new ExcelService(_logger);
        var nonExistentPath = "/tmp/non-existent-file.xlsx";

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => service.ReadFileInfoAsync(nonExistentPath));
    }
}
