using System.Xml;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates xl/worksheets/_rels/sheet{N}.xml.rels for worksheets that have
/// tables, embedded charts, or embedded images.
/// </summary>
internal static class XlsxWorksheetRelsPart
{
    private const string Ns          = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string TableType   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    private const string DrawingType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

    public static byte[] Generate(IReadOnlyList<XlsxTableInfo> tables,
                                   int drawingId      = 0,
                                   int drawingLocalRId = 0)
    {
        using var ms  = new System.IO.MemoryStream();
        using var xml = XmlWriter.Create(ms, new XmlWriterSettings
        {
            Encoding           = new System.Text.UTF8Encoding(false),
            Indent             = false,
            OmitXmlDeclaration = false,
        });

        xml.WriteStartDocument(true);
        xml.WriteStartElement("Relationships", Ns);

        for (int i = 0; i < tables.Count; i++)
        {
            xml.WriteStartElement("Relationship", Ns);
            xml.WriteAttributeString("Id",     $"rId{i + 1}");
            xml.WriteAttributeString("Type",   TableType);
            xml.WriteAttributeString("Target", $"../tables/table{tables[i].TableId}.xml");
            xml.WriteEndElement();
        }

        if (drawingId > 0 && drawingLocalRId > 0)
        {
            xml.WriteStartElement("Relationship", Ns);
            xml.WriteAttributeString("Id",     $"rId{drawingLocalRId}");
            xml.WriteAttributeString("Type",   DrawingType);
            xml.WriteAttributeString("Target", $"../drawings/drawing{drawingId}.xml");
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        xml.Flush();
        return ms.ToArray();
    }
}
