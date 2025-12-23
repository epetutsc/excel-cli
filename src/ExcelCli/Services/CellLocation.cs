namespace ExcelCli.Services;

/// <summary>
/// Location of a cell
/// </summary>
public record CellLocation(string SheetName, string CellAddress, string Value);
