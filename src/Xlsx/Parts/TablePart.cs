using System.Globalization;
using System.Xml;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates xl/tables/table{N}.xml for one Excel table.
/// </summary>
internal static class XlsxTablePart
{
    private const string Ns    = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NsMc  = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string NsXr  = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private const string NsXr3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3";

    public static byte[] Generate(XlsxTableInfo table)
    {
        var def = table.Definition;

        // autoFilter covers the table minus the totals row
        var autoFilterRange = def.HasTotalsRow ? ExcludeLastRow(table.Range) : table.Range;

        using var ms  = new System.IO.MemoryStream();
        using var xml = XmlWriter.Create(ms, new XmlWriterSettings
        {
            Encoding           = new System.Text.UTF8Encoding(false),
            Indent             = false,
            OmitXmlDeclaration = false,
        });

        xml.WriteStartDocument(true);

        xml.WriteStartElement("table", Ns);
        xml.WriteAttributeString("xmlns", "mc",  null, NsMc);
        xml.WriteAttributeString("mc",    "Ignorable", NsMc, "xr xr3");
        xml.WriteAttributeString("xmlns", "xr",  null, NsXr);
        xml.WriteAttributeString("xmlns", "xr3", null, NsXr3);
        xml.WriteAttributeString("id",          table.TableId.ToString(CultureInfo.InvariantCulture));
        xml.WriteAttributeString("name",        def.Name);
        xml.WriteAttributeString("displayName", def.DisplayName);
        xml.WriteAttributeString("ref",         table.Range);
        if (def.HasTotalsRow)
            xml.WriteAttributeString("totalsRowCount", "1");

        // <autoFilter>
        xml.WriteStartElement("autoFilter", Ns);
        xml.WriteAttributeString("ref", autoFilterRange);
        xml.WriteEndElement();

        // <tableColumns>
        xml.WriteStartElement("tableColumns", Ns);
        xml.WriteAttributeString("count",
            def.Columns.Count.ToString(CultureInfo.InvariantCulture));

        for (int i = 0; i < def.Columns.Count; i++)
        {
            var col = def.Columns[i];
            xml.WriteStartElement("tableColumn", Ns);
            xml.WriteAttributeString("id",   (i + 1).ToString(CultureInfo.InvariantCulture));
            xml.WriteAttributeString("name", col.Name);
            if (col.TotalsFunction != XlsxTotalsRowFunction.None)
                xml.WriteAttributeString("totalsRowFunction", TotalsXml(col.TotalsFunction));
            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // tableColumns

        // <tableStyleInfo>
        xml.WriteStartElement("tableStyleInfo", Ns);
        xml.WriteAttributeString("name",              def.StyleName);
        xml.WriteAttributeString("showFirstColumn",   def.ShowFirstColumn ? "1" : "0");
        xml.WriteAttributeString("showLastColumn",    def.ShowLastColumn  ? "1" : "0");
        xml.WriteAttributeString("showRowStripes",    def.ShowRowStripes  ? "1" : "0");
        xml.WriteAttributeString("showColumnStripes", def.ShowColStripes  ? "1" : "0");
        xml.WriteEndElement();

        xml.WriteEndElement(); // table
        xml.Flush();
        return ms.ToArray();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string ExcludeLastRow(string range)
    {
        var parts = range.Split(':');
        if (parts.Length != 2) return range;
        var (endRow, endCol) = CellReference.Parse(parts[1]);
        var (startRow, _)    = CellReference.Parse(parts[0]);
        if (endRow <= startRow) return range;
        return $"{parts[0]}:{CellReference.FromRowCol(endRow - 1, endCol)}";
    }

    private static string TotalsXml(XlsxTotalsRowFunction f) => f switch
    {
        XlsxTotalsRowFunction.Sum       => "sum",
        XlsxTotalsRowFunction.Average   => "average",
        XlsxTotalsRowFunction.Count     => "count",
        XlsxTotalsRowFunction.CountNums => "countNums",
        XlsxTotalsRowFunction.Max       => "max",
        XlsxTotalsRowFunction.Min       => "min",
        XlsxTotalsRowFunction.StdDev    => "stdDev",
        XlsxTotalsRowFunction.Var       => "var",
        _                           => string.Empty,
    };
}
