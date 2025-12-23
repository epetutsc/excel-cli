using ExcelCli.Services;
using NSubstitute;
using Serilog;
using System.IO.Abstractions.TestingHelpers;
using Xunit;

namespace ExcelCli.Tests;

/// <summary>
/// Tests for ExcelService constructor validation
/// </summary>
public class ExcelServiceConstructorTests
{
    private readonly ILogger _logger;
    private readonly MockFileSystem _fileSystem;

    public ExcelServiceConstructorTests()
    {
        _logger = Substitute.For<ILogger>();
        _fileSystem = new MockFileSystem();
    }

    [Fact]
    public void Constructor_WithNullLogger_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => new ExcelService(null!, _fileSystem));
    }

    [Fact]
    public void Constructor_WithNullFileSystem_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => new ExcelService(_logger, null!));
    }

    [Fact]
    public void Constructor_WithValidArguments_CreatesInstance()
    {
        var service = new ExcelService(_logger, _fileSystem);
        Assert.NotNull(service);
    }
}
