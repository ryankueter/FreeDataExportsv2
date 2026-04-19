using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace FreeDataExportsv2.Internal;

internal static class XlsxXmlHelper
{
    /// <summary>
    /// Serialises an <see cref="XDocument"/> to a UTF-8 byte array
    /// (with XML declaration, no BOM).
    /// </summary>
    public static byte[] ToXmlBytes(XDocument doc)
    {
        using var ms  = new System.IO.MemoryStream();
        using var xml = XmlWriter.Create(ms, new XmlWriterSettings
        {
            Encoding           = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent             = false,
            OmitXmlDeclaration = false,
        });
        doc.WriteTo(xml);
        xml.Flush();
        return ms.ToArray();
    }

    /// <summary>Formats a double in a way that roundtrips without scientific notation.</summary>
    public static string F(double v) =>
        v.ToString("G", CultureInfo.InvariantCulture);
}
