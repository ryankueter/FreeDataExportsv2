using System.Text;

namespace FreeDataExportsv2.Internal;

/// <summary>Generates settings.xml for an ODS package (view/configuration settings).</summary>
internal static class OdsSettingsPart
{
    public static byte[] Generate(XlsxWorkbook workbook)
    {
        var sb = new StringBuilder();
        sb.Append("""
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-settings
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
              office:version="1.4">
              <office:settings>
                <config:config-item-set config:name="ooo:view-settings">
                  <config:config-item config:name="ViewStyleName" config:type="string">Default</config:config-item>
                  <config:config-item-map-indexed config:name="Views">
                    <config:config-item-map-entry>
                      <config:config-item config:name="ViewId" config:type="string">view1</config:config-item>
                      <config:config-item config:name="ActiveTable" config:type="string">
            """);
        sb.Append(XmlEsc(workbook.Sheets.Count > 0 ? workbook.Sheets[0].Name : "Sheet1"));
        sb.Append("""
                      </config:config-item>
                      <config:config-item config:name="ZoomValue" config:type="int">100</config:config-item>
                      <config:config-item config:name="PageViewZoomValue" config:type="int">60</config:config-item>
                      <config:config-item config:name="ShowPageBreakPreview" config:type="boolean">false</config:config-item>
                    </config:config-item-map-entry>
                  </config:config-item-map-indexed>
                </config:config-item-set>
                <config:config-item-set config:name="ooo:configuration-settings">
                  <config:config-item config:name="IsDocumentShared" config:type="boolean">false</config:config-item>
                  <config:config-item config:name="SaveVersionOnClose" config:type="boolean">false</config:config-item>
                </config:config-item-set>
              </office:settings>
            </office:document-settings>
            """);

        return Encoding.UTF8.GetBytes(sb.ToString());
    }

    private static string XmlEsc(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
