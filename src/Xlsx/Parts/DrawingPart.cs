using System.Globalization;
using System.Xml;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates xl/drawings/drawing{N}.xml — the spreadsheet-drawing container that
/// anchors charts and/or images to a worksheet.
/// </summary>
internal static class XlsxDrawingPart
{
    private const string NsXdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string NsA   = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string NsC   = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string NsR   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static byte[] Generate(IReadOnlyList<ChartInfo>  charts,
                                   IReadOnlyList<XlsxImageEntry> images)
    {
        using var ms  = new System.IO.MemoryStream();
        using var xml = XmlWriter.Create(ms, new XmlWriterSettings
        {
            Encoding           = new System.Text.UTF8Encoding(false),
            Indent             = false,
            OmitXmlDeclaration = false,
        });

        xml.WriteStartDocument(true);

        // <xdr:wsDr>
        xml.WriteStartElement("xdr", "wsDr", NsXdr);
        xml.WriteAttributeString("xmlns", "xdr", null, NsXdr);
        xml.WriteAttributeString("xmlns", "a",   null, NsA);
        xml.WriteAttributeString("xmlns", "c",   null, NsC);
        xml.WriteAttributeString("xmlns", "r",   null, NsR);

        // ── Charts ────────────────────────────────────────────────────────────
        for (int i = 0; i < charts.Count; i++)
        {
            var chart  = charts[i];
            var anchor = chart.Anchor;
            string rId = $"rId{i + 1}";

            xml.WriteStartElement("xdr", "twoCellAnchor", NsXdr);
            xml.WriteAttributeString("editAs", "oneCell");

            WriteAnchorPos(xml, anchor.FromCol, anchor.FromColOff, anchor.FromRow, anchor.FromRowOff, "from");
            WriteAnchorPos(xml, anchor.ToCol,   anchor.ToColOff,   anchor.ToRow,   anchor.ToRowOff,   "to");

            // <xdr:graphicFrame>
            xml.WriteStartElement("xdr", "graphicFrame", NsXdr);
            xml.WriteAttributeString("macro", "");

            xml.WriteStartElement("xdr", "nvGraphicFramePr", NsXdr);
            xml.WriteStartElement("xdr", "cNvPr", NsXdr);
            xml.WriteAttributeString("id",   I(i + 2));
            xml.WriteAttributeString("name", $"Chart {i + 1}");
            xml.WriteEndElement(); // cNvPr
            xml.WriteStartElement("xdr", "cNvGraphicFramePr", NsXdr);
            xml.WriteEndElement();
            xml.WriteEndElement(); // nvGraphicFramePr

            xml.WriteStartElement("xdr", "xfrm", NsXdr);
            xml.WriteStartElement("a", "off", NsA);
            xml.WriteAttributeString("x", "0"); xml.WriteAttributeString("y", "0");
            xml.WriteEndElement();
            xml.WriteStartElement("a", "ext", NsA);
            xml.WriteAttributeString("cx", "0"); xml.WriteAttributeString("cy", "0");
            xml.WriteEndElement();
            xml.WriteEndElement(); // xfrm

            xml.WriteStartElement("a", "graphic", NsA);
            xml.WriteStartElement("a", "graphicData", NsA);
            xml.WriteAttributeString("uri", NsC);
            xml.WriteStartElement("c", "chart", NsC);
            xml.WriteAttributeString("r", "id", NsR, rId);
            xml.WriteEndElement(); // c:chart
            xml.WriteEndElement(); // a:graphicData
            xml.WriteEndElement(); // a:graphic

            xml.WriteEndElement(); // xdr:graphicFrame

            xml.WriteStartElement("xdr", "clientData", NsXdr);
            xml.WriteEndElement();

            xml.WriteEndElement(); // xdr:twoCellAnchor
        }

        // ── Images ────────────────────────────────────────────────────────────
        int shapeIdBase = charts.Count + 2; // shape IDs: 2..N for charts, then continue
        for (int i = 0; i < images.Count; i++)
        {
            var img    = images[i];
            var anchor = img.Anchor;
            string rId = $"rId{img.DrawingRId}";
            int    shapeId = shapeIdBase + i;

            xml.WriteStartElement("xdr", "twoCellAnchor", NsXdr);
            xml.WriteAttributeString("editAs", "oneCell");

            WriteAnchorPos(xml, anchor.FromCol, anchor.FromColOff, anchor.FromRow, anchor.FromRowOff, "from");
            WriteAnchorPos(xml, anchor.ToCol,   anchor.ToColOff,   anchor.ToRow,   anchor.ToRowOff,   "to");

            // <xdr:pic>
            xml.WriteStartElement("xdr", "pic", NsXdr);

            // <xdr:nvPicPr>
            xml.WriteStartElement("xdr", "nvPicPr", NsXdr);
            xml.WriteStartElement("xdr", "cNvPr", NsXdr);
            xml.WriteAttributeString("id",   I(shapeId));
            xml.WriteAttributeString("name", $"Image {i + 1}");
            xml.WriteEndElement(); // cNvPr
            xml.WriteStartElement("xdr", "cNvPicPr", NsXdr);
            xml.WriteStartElement("a", "picLocks", NsA);
            xml.WriteAttributeString("noChangeAspect", "1");
            xml.WriteEndElement(); // a:picLocks
            xml.WriteEndElement(); // cNvPicPr
            xml.WriteEndElement(); // nvPicPr

            // <xdr:blipFill>
            xml.WriteStartElement("xdr", "blipFill", NsXdr);
            xml.WriteStartElement("a", "blip", NsA);
            xml.WriteAttributeString("r", "embed", NsR, rId);
            xml.WriteAttributeString("cstate", "print");
            xml.WriteEndElement(); // a:blip
            xml.WriteStartElement("a", "stretch", NsA);
            xml.WriteStartElement("a", "fillRect", NsA);
            xml.WriteEndElement(); // a:fillRect
            xml.WriteEndElement(); // a:stretch
            xml.WriteEndElement(); // blipFill

            // <xdr:spPr>
            xml.WriteStartElement("xdr", "spPr", NsXdr);
            xml.WriteStartElement("a", "xfrm", NsA);
            xml.WriteStartElement("a", "off", NsA);
            xml.WriteAttributeString("x", "0"); xml.WriteAttributeString("y", "0");
            xml.WriteEndElement(); // a:off
            xml.WriteStartElement("a", "ext", NsA);
            xml.WriteAttributeString("cx", "0"); xml.WriteAttributeString("cy", "0");
            xml.WriteEndElement(); // a:ext
            xml.WriteEndElement(); // a:xfrm
            xml.WriteStartElement("a", "prstGeom", NsA);
            xml.WriteAttributeString("prst", "rect");
            xml.WriteStartElement("a", "avLst", NsA);
            xml.WriteEndElement(); // a:avLst
            xml.WriteEndElement(); // a:prstGeom
            xml.WriteEndElement(); // xdr:spPr

            xml.WriteEndElement(); // xdr:pic

            xml.WriteStartElement("xdr", "clientData", NsXdr);
            xml.WriteEndElement();

            xml.WriteEndElement(); // xdr:twoCellAnchor
        }

        xml.WriteEndElement(); // xdr:wsDr
        xml.Flush();
        return ms.ToArray();
    }

    private static void WriteAnchorPos(XmlWriter xml, int col, int colOff, int row, int rowOff, string tag)
    {
        xml.WriteStartElement("xdr", tag, NsXdr);
        E(xml, "col",    I(col));
        E(xml, "colOff", I(colOff));
        E(xml, "row",    I(row));
        E(xml, "rowOff", I(rowOff));
        xml.WriteEndElement();
    }

    private static void E(XmlWriter xml, string local, string value)
    {
        xml.WriteStartElement("xdr", local, NsXdr);
        xml.WriteString(value);
        xml.WriteEndElement();
    }

    private static string I(int v) => v.ToString(CultureInfo.InvariantCulture);
}
