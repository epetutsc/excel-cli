namespace ExcelCli.Services;

/// <summary>
/// Information about a worksheet
/// </summary>
public record SheetInfo(string Name, int RowCount, int ColumnCount);
