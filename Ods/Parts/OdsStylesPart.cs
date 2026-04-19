using System.Text;

namespace FreeDataExportsv2.Internal;

/// <summary>Generates styles.xml for an ODS package (named/page styles).</summary>
internal static class OdsStylesPart
{
    public static byte[] Generate()
    {
        // Minimal styles.xml — cell automatic styles live in content.xml.
        // We provide only the named styles (Default, Heading, etc.) and master pages
        // so that LibreOffice and other readers find the expected structure.
        const string xml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-styles
              xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
              xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
              xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
              xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
              xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
              xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
              xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
              xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"
              office:version="1.4">
              <office:font-face-decls>
                <style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'" style:font-family-generic="swiss" style:font-pitch="variable"/>
                <style:font-face style:name="Calibri" svg:font-family="Calibri" style:font-family-generic="swiss"/>
              </office:font-face-decls>
              <office:styles>
                <style:default-style style:family="table-cell">
                  <style:paragraph-properties style:tab-stop-distance="0.5in"/>
                  <style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt" fo:language="en" fo:country="US"/>
                </style:default-style>
                <style:default-style style:family="graphic">
                  <style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf"/>
                </style:default-style>
                <style:style style:name="Default" style:family="table-cell"/>
                <style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default">
                  <style:text-properties fo:font-size="24pt" fo:font-weight="bold"/>
                </style:style>
                <style:style style:name="Text" style:family="table-cell" style:parent-style-name="Default"/>
                <style:style style:name="Note" style:family="table-cell" style:parent-style-name="Text">
                  <style:table-cell-properties fo:background-color="#ffffcc" fo:border="0.74pt solid #808080"/>
                  <style:text-properties fo:color="#333333"/>
                </style:style>
                <style:style style:name="Good" style:family="table-cell" style:parent-style-name="Default">
                  <style:table-cell-properties fo:background-color="#ccffcc"/>
                  <style:text-properties fo:color="#006600"/>
                </style:style>
                <style:style style:name="Bad" style:family="table-cell" style:parent-style-name="Default">
                  <style:table-cell-properties fo:background-color="#ffcccc"/>
                  <style:text-properties fo:color="#cc0000"/>
                </style:style>
                <number:number-style style:name="N0">
                  <number:number number:min-integer-digits="1"/>
                </number:number-style>
                <draw:marker draw:name="Arrowheads_20_1" draw:display-name="Arrowheads 1" svg:viewBox="0 0 20 30" svg:d="M10 0l-10 30h20z"/>
              </office:styles>
              <office:automatic-styles>
                <style:page-layout style:name="Mpm1">
                  <style:page-layout-properties style:writing-mode="lr-tb"/>
                  <style:header-style>
                    <style:header-footer-properties fo:min-height="0.2953in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-bottom="0.0984in"/>
                  </style:header-style>
                  <style:footer-style>
                    <style:header-footer-properties fo:min-height="0.2953in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-top="0.0984in"/>
                  </style:footer-style>
                </style:page-layout>
              </office:automatic-styles>
              <office:master-styles>
                <style:master-page style:name="Default" style:page-layout-name="Mpm1">
                  <style:header><text:p><text:sheet-name>???</text:sheet-name></text:p></style:header>
                  <style:header-left style:display="false"/>
                  <style:footer><text:p>Page <text:page-number>1</text:page-number></text:p></style:footer>
                  <style:footer-left style:display="false"/>
                </style:master-page>
              </office:master-styles>
            </office:document-styles>
            """;

        return Encoding.UTF8.GetBytes(xml);
    }
}
