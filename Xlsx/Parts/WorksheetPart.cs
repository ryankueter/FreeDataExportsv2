using System.Globalization;
using System.Xml;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

internal static class XlsxWorksheetPart
{
    private const string NsMain  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NsR     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string NsMc    = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string NsX14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private const string NsXr    = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private const string NsXr2   = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
    private const string NsXr3   = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3";

    public static byte[] Generate(XlsxWorksheet sheet)
    {
        using var ms  = new System.IO.MemoryStream();
        using var xml = XmlWriter.Create(ms, new XmlWriterSettings
        {
            Encoding           = new System.Text.UTF8Encoding(false),
            Indent             = false,
            OmitXmlDeclaration = false,
        });

        xml.WriteStartDocument(true);

        xml.WriteStartElement("worksheet", NsMain);
        xml.WriteAttributeString("xmlns", "r",     null, NsR);
        xml.WriteAttributeString("xmlns", "mc",    null, NsMc);
        xml.WriteAttributeString("mc",    "Ignorable", NsMc, "x14ac xr xr2 xr3");
        xml.WriteAttributeString("xmlns", "x14ac", null, NsX14ac);
        xml.WriteAttributeString("xmlns", "xr",    null, NsXr);
        xml.WriteAttributeString("xmlns", "xr2",   null, NsXr2);
        xml.WriteAttributeString("xmlns", "xr3",   null, NsXr3);

        if (sheet.TabColor is not null)
        {
            xml.WriteStartElement("sheetPr",  NsMain);
            xml.WriteStartElement("tabColor", NsMain);
            xml.WriteAttributeString("rgb", sheet.TabColor);
            xml.WriteEndElement();
            xml.WriteEndElement();
        }

        xml.WriteStartElement("dimension", NsMain);
        xml.WriteAttributeString("ref", sheet.DimensionRef ?? ComputeDimensionRef(sheet));
        xml.WriteEndElement();

        xml.WriteStartElement("sheetViews", NsMain);
        xml.WriteStartElement("sheetView",  NsMain);
        if (sheet.IsTabSelected)
            xml.WriteAttributeString("tabSelected", "1");
        xml.WriteAttributeString("workbookViewId", "0");
        xml.WriteEndElement();
        xml.WriteEndElement();

        xml.WriteStartElement("sheetFormatPr", NsMain);
        xml.WriteAttributeString("defaultRowHeight", F(sheet.DefaultRowHeight));
        xml.WriteAttributeString("dyDescent", NsX14ac, F(sheet.DyDescent));
        xml.WriteEndElement();

        int? pageStyleIdx = sheet.GetSheetStyleIndex();
        string? pageStyleStr = pageStyleIdx.HasValue
            ? pageStyleIdx.Value.ToString(CultureInfo.InvariantCulture)
            : null;

        if (sheet.ColWidths.Count > 0 || pageStyleIdx.HasValue)
        {
            xml.WriteStartElement("cols", NsMain);

            if (sheet.ColWidths.Count > 0)
            {
                int highestCustomCol = 0;
                foreach (var (min, max, width) in sheet.ColWidths)
                {
                    xml.WriteStartElement("col", NsMain);
                    xml.WriteAttributeString("min",         min.ToString(CultureInfo.InvariantCulture));
                    xml.WriteAttributeString("max",         max.ToString(CultureInfo.InvariantCulture));
                    xml.WriteAttributeString("width",       F(width));
                    xml.WriteAttributeString("customWidth", "1");
                    if (pageStyleStr is not null)
                        xml.WriteAttributeString("style", pageStyleStr);
                    xml.WriteEndElement();
                    if (max > highestCustomCol) highestCustomCol = max;
                }
                if (pageStyleIdx.HasValue && highestCustomCol < 16384)
                {
                    // Include width + customWidth so Excel reliably applies the style attribute.
                    // Without customWidth="1", Excel silently ignores style on col elements.
                    xml.WriteStartElement("col", NsMain);
                    xml.WriteAttributeString("min",         (highestCustomCol + 1).ToString(CultureInfo.InvariantCulture));
                    xml.WriteAttributeString("max",         "16384");
                    xml.WriteAttributeString("width",       "8.43");
                    xml.WriteAttributeString("customWidth", "1");
                    xml.WriteAttributeString("style",       pageStyleStr!);
                    xml.WriteEndElement();
                }
            }
            else if (pageStyleIdx.HasValue)
            {
                xml.WriteStartElement("col", NsMain);
                xml.WriteAttributeString("min",         "1");
                xml.WriteAttributeString("max",         "16384");
                xml.WriteAttributeString("width",       "8.43");
                xml.WriteAttributeString("customWidth", "1");
                xml.WriteAttributeString("style",       pageStyleStr!);
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        xml.WriteStartElement("sheetData", NsMain);
        foreach (var row in sheet.Rows)
        {
            xml.WriteStartElement("row", NsMain);
            xml.WriteAttributeString("r", row.RowIndex.ToString(CultureInfo.InvariantCulture));
            if (pageStyleStr is not null)
            {
                xml.WriteAttributeString("s", pageStyleStr);
                xml.WriteAttributeString("customFormat", "1");
            }
            foreach (var cell in row.Cells)
                WriteCell(xml, cell);
            xml.WriteEndElement();
        }
        xml.WriteEndElement();

        var m = sheet.XlsxPageMargins;
        xml.WriteStartElement("pageMargins", NsMain);
        xml.WriteAttributeString("left",   F(m.Left));
        xml.WriteAttributeString("right",  F(m.Right));
        xml.WriteAttributeString("top",    F(m.Top));
        xml.WriteAttributeString("bottom", F(m.Bottom));
        xml.WriteAttributeString("header", F(m.Header));
        xml.WriteAttributeString("footer", F(m.Footer));
        xml.WriteEndElement();

        if (sheet.DrawingLocalRId > 0)
        {
            xml.WriteStartElement("drawing", NsMain);
            xml.WriteAttributeString("r", "id", NsR, $"rId{sheet.DrawingLocalRId}");
            xml.WriteEndElement();
        }

        if (sheet.Tables.Count > 0)
        {
            xml.WriteStartElement("tableParts", NsMain);
            xml.WriteAttributeString("count",
                sheet.Tables.Count.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < sheet.Tables.Count; i++)
            {
                xml.WriteStartElement("tablePart", NsMain);
                xml.WriteAttributeString("r", "id", NsR, $"rId{i + 1}");
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.Flush();
        return ms.ToArray();
    }

    private static void WriteCell(XmlWriter xml, XlsxCell cell)
    {
        xml.WriteStartElement("c", NsMain);
        xml.WriteAttributeString("r", cell.Reference);
        if (cell.StyleIndex.HasValue)
            xml.WriteAttributeString("s", cell.StyleIndex.Value
                .ToString(CultureInfo.InvariantCulture));

        switch (cell.Value)
        {
            case CellValue.Number n:
                xml.WriteElementString("v", NsMain,
                    n.Value.ToString("R", CultureInfo.InvariantCulture));
                break;

            case CellValue.Text t:
                xml.WriteAttributeString("t", "inlineStr");
                xml.WriteStartElement("is", NsMain);
                xml.WriteElementString("t", NsMain, t.Value);
                xml.WriteEndElement();
                break;

            case CellValue.Boolean b:
                xml.WriteAttributeString("t", "b");
                xml.WriteElementString("v", NsMain, b.Value ? "1" : "0");
                break;

            case CellValue.Date d:
                xml.WriteElementString("v", NsMain,
                    d.Value.ToOADate().ToString("R", CultureInfo.InvariantCulture));
                break;

            case CellValue.Formula f:
                if (f.ResultType != FormulaResultType.Number)
                    xml.WriteAttributeString("t", FormulaTypeCode(f.ResultType));
                xml.WriteElementString("f", NsMain, f.Expression);
                if (f.CachedResult is not null)
                    xml.WriteElementString("v", NsMain, CachedResultStr(f.CachedResult));
                break;

            case CellValue.Error e:
                xml.WriteAttributeString("t", "e");
                xml.WriteElementString("v", NsMain, e.Code.ToXmlString());
                break;

            case null:
                break;
        }

        xml.WriteEndElement();
    }

    private static string FormulaTypeCode(FormulaResultType t) => t switch
    {
        FormulaResultType.Text    => "str",
        FormulaResultType.Boolean => "b",
        FormulaResultType.Error   => "e",
        _                         => "n",
    };

    private static string CachedResultStr(object r) => r switch
    {
        bool   b => b ? "1" : "0",
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        float  f => f.ToString("R", CultureInfo.InvariantCulture),
        _        => r.ToString() ?? string.Empty,
    };

    private static string ComputeDimensionRef(XlsxWorksheet sheet)
    {
        var refs = sheet.Rows
            .SelectMany(r => r.Cells)
            .Select(c => c.Reference)
            .Where(r => !string.IsNullOrEmpty(r))
            .Select(CellReference.Parse)
            .ToList();

        if (refs.Count == 0) return "A1";
        int minRow = refs.Min(p => p.row), maxRow = refs.Max(p => p.row);
        int minCol = refs.Min(p => p.col), maxCol = refs.Max(p => p.col);
        var tl = CellReference.FromRowCol(minRow, minCol);
        var br = CellReference.FromRowCol(maxRow, maxCol);
        return tl == br ? tl : $"{tl}:{br}";
    }

    private static string F(double v) =>
        v.ToString("G", CultureInfo.InvariantCulture);
}
