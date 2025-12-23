namespace ExcelCli.Services;

/// <summary>
/// Information about an Excel file
/// </summary>
public record FileInfo(string FileName, long FileSize, DateTime LastModified, int SheetCount);
