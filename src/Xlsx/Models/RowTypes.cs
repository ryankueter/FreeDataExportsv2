namespace FreeDataExportsv2.Internal;

internal class XlsxPageMargins
{
    public double Left   { get; set; } = 0.7;
    public double Right  { get; set; } = 0.7;
    public double Top    { get; set; } = 0.75;
    public double Bottom { get; set; } = 0.75;
    public double Header { get; set; } = 0.3;
    public double Footer { get; set; } = 0.3;
}

internal class XlsxRow
{
    public int             RowIndex { get; set; }
    public List<XlsxCell> Cells    { get; } = [];
}

internal class XlsxCell
{
    public string     Reference  { get; set; } = string.Empty;
    public FreeDataExportsv2.CellValue? Value { get; set; }
    public int?       StyleIndex { get; set; }
}

/// <summary>Extension methods on <see cref="FreeDataExportsv2.ErrorCode"/>.</summary>
internal static class XlsxErrorCodeExtensions
{
    public static string ToXmlString(this FreeDataExportsv2.ErrorCode code) => code switch
    {
        FreeDataExportsv2.ErrorCode.DivisionByZero   => "#DIV/0!",
        FreeDataExportsv2.ErrorCode.NotAvailable     => "#N/A",
        FreeDataExportsv2.ErrorCode.InvalidName      => "#NAME?",
        FreeDataExportsv2.ErrorCode.NullIntersection => "#NULL!",
        FreeDataExportsv2.ErrorCode.InvalidNumber    => "#NUM!",
        FreeDataExportsv2.ErrorCode.InvalidReference => "#REF!",
        FreeDataExportsv2.ErrorCode.InvalidValue     => "#VALUE!",
        _                                            => "#VALUE!",
    };

    /// <summary>Returns an ODS formula string that evaluates to this error code.</summary>
    public static string ToOdsFormula(this FreeDataExportsv2.ErrorCode code) => code switch
    {
        FreeDataExportsv2.ErrorCode.DivisionByZero   => "of:=1/0",
        FreeDataExportsv2.ErrorCode.NotAvailable     => "of:=NA()",
        FreeDataExportsv2.ErrorCode.InvalidName      => "of:=ERR:522()",
        FreeDataExportsv2.ErrorCode.NullIntersection => "of:=ERR:521()",
        FreeDataExportsv2.ErrorCode.InvalidNumber    => "of:=SQRT(-1)",
        FreeDataExportsv2.ErrorCode.InvalidReference => "of:=ERR:524()",
        FreeDataExportsv2.ErrorCode.InvalidValue     => "of:=ERR:519()",
        _                                            => "of:=ERR:519()",
    };
}
