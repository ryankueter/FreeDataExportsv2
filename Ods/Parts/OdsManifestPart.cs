using System.Text;

namespace FreeDataExportsv2.Internal;

/// <summary>Generates META-INF/manifest.xml for an ODS package.</summary>
internal static class OdsManifestPart
{
    public static byte[] Generate(XlsxWorkbook workbook, bool hasImages)
    {
        var sb = new StringBuilder();
        sb.Append("""
            <?xml version="1.0" encoding="UTF-8"?>
            <manifest:manifest
              xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
              manifest:version="1.4">
              <manifest:file-entry manifest:full-path="/" manifest:version="1.4" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
              <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
              <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
              <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
              <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>

            """);

        // Image entries
        var seenExts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in workbook.Sheets)
        {
            foreach (var img in sheet.Images)
            {
                string path = $"Pictures/image{img.MediaId}.{img.Extension}";
                sb.Append($"  <manifest:file-entry manifest:full-path=\"{path}\" manifest:media-type=\"{img.ContentType}\"/>\n");
            }
        }

        // Chart object entries
        foreach (var sheet in workbook.Sheets)
        {
            foreach (var c in sheet.Charts)
            {
                int id = c.ChartId;
                sb.Append($"  <manifest:file-entry manifest:full-path=\"Object {id}/\" manifest:media-type=\"application/vnd.oasis.opendocument.chart\"/>\n");
                sb.Append($"  <manifest:file-entry manifest:full-path=\"Object {id}/content.xml\" manifest:media-type=\"text/xml\"/>\n");
                sb.Append($"  <manifest:file-entry manifest:full-path=\"Object {id}/styles.xml\" manifest:media-type=\"text/xml\"/>\n");
                sb.Append($"  <manifest:file-entry manifest:full-path=\"Object {id}/meta.xml\" manifest:media-type=\"text/xml\"/>\n");
            }
        }

        sb.Append("</manifest:manifest>\n");
        return Encoding.UTF8.GetBytes(sb.ToString());
    }
}
