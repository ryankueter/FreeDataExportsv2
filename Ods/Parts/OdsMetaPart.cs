using System.Text;

namespace FreeDataExportsv2.Internal;

/// <summary>Generates meta.xml for an ODS package.</summary>
internal static class OdsMetaPart
{
    public static byte[] Generate(OdsFile file)
    {
        string created  = file.Created.ToString("yyyy-MM-ddTHH:mm:ssZ");
        string modified = file.Modified.ToString("yyyy-MM-ddTHH:mm:ssZ");
        string creator  = XmlEsc(file.Creator);

        var xml = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-meta
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
              xmlns:dc="http://purl.org/dc/elements/1.1/"
              office:version="1.4">
              <office:meta>
                <meta:generator>FreeDataExportsv2</meta:generator>
                <dc:creator>{creator}</dc:creator>
                <dc:date>{modified}</dc:date>
                <meta:creation-date>{created}</meta:creation-date>
                <meta:editing-cycles>1</meta:editing-cycles>
              </office:meta>
            </office:document-meta>
            """;

        return Encoding.UTF8.GetBytes(xml);
    }

    private static string XmlEsc(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
