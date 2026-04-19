using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

internal sealed class XlsxWorkbook
{
    public List<XlsxWorksheet> Sheets { get; } = [];
    public XlsxStyles          XlsxStyles { get; } = new();
}
