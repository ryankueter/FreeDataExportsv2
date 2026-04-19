namespace FreeDataExportsv2.Internal;

/// <summary>Captures cell-level errors that occur during workbook construction.</summary>
internal sealed class XlsxErrorLog
{
    private readonly List<XlsxErrorRecord> _records = [];

    public IReadOnlyList<XlsxErrorRecord> Records => _records;

    public void Add(string? sheetName, string? cellRef, object? attemptedValue, Exception ex)
    {
        _records.Add(new XlsxErrorRecord(sheetName, cellRef, attemptedValue,
                                     ex.GetType().Name, ex.Message));
    }
}

internal sealed class XlsxErrorRecord
{
    public string?  SheetName      { get; }
    public string?  CellRef        { get; }
    public object?  AttemptedValue { get; }
    public string   ExceptionType  { get; }
    public string   ErrorMessage   { get; }

    public XlsxErrorRecord(string? sheetName, string? cellRef, object? attemptedValue,
                       string exceptionType, string errorMessage)
    {
        SheetName      = sheetName;
        CellRef        = cellRef;
        AttemptedValue = attemptedValue;
        ExceptionType  = exceptionType;
        ErrorMessage   = errorMessage;
    }
}
