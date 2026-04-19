using System.Xml;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates xl/drawings/_rels/drawing{N}.xml.rels — the relationships file that
/// links each chart / image anchor in a drawing to its target file.
/// </summary>
internal static class XlsxDrawingRelsPart
{
    private const string Ns        = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ChartType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    private const string ImageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

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
        xml.WriteStartElement("Relationships", Ns);

        // Charts — rId1 … rId{N}
        for (int i = 0; i < charts.Count; i++)
        {
            xml.WriteStartElement("Relationship", Ns);
            xml.WriteAttributeString("Id",     $"rId{i + 1}");
            xml.WriteAttributeString("Type",   ChartType);
            xml.WriteAttributeString("Target", $"../charts/chart{charts[i].ChartId}.xml");
            xml.WriteEndElement();
        }

        // Images — rId{N+1} … rId{N+M}  (DrawingRId is pre-assigned by XlsxFile)
        foreach (var img in images)
        {
            xml.WriteStartElement("Relationship", Ns);
            xml.WriteAttributeString("Id",     $"rId{img.DrawingRId}");
            xml.WriteAttributeString("Type",   ImageType);
            xml.WriteAttributeString("Target", $"../media/image{img.MediaId}.{img.Extension}");
            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // Relationships
        xml.Flush();
        return ms.ToArray();
    }
}
