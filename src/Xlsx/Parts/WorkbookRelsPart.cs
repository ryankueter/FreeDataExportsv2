using System.Globalization;
using System.Xml.Linq;

namespace FreeDataExportsv2.Internal;

/// <summary>Generates xl/_rels/workbook.xml.rels.</summary>
internal static class XlsxWorkbookRelsPart
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string SheetType  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    private const string StylesType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    private const string ThemeType  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    public static byte[] Generate(XlsxWorkbook workbook)
    {
        var rels = new List<XElement>();

        for (int i = 0; i < workbook.Sheets.Count; i++)
        {
            rels.Add(new XElement(Ns + "Relationship",
                new XAttribute("Id",     $"rId{i + 1}"),
                new XAttribute("Type",   SheetType),
                new XAttribute("Target", $"worksheets/sheet{i + 1}.xml")));
        }

        int nextId = workbook.Sheets.Count + 1;

        rels.Add(new XElement(Ns + "Relationship",
            new XAttribute("Id",     $"rId{nextId++}"),
            new XAttribute("Type",   StylesType),
            new XAttribute("Target", "styles.xml")));

        rels.Add(new XElement(Ns + "Relationship",
            new XAttribute("Id",     $"rId{nextId}"),
            new XAttribute("Type",   ThemeType),
            new XAttribute("Target", "theme/theme1.xml")));

        var root = new XElement(Ns + "Relationships", rels);
        return XlsxXmlHelper.ToXmlBytes(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }
}
