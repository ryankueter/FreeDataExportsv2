using System.Xml.Linq;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates docProps/app.xml — extended application properties.
/// HeadingPairs and TitlesOfParts are derived from the workbook's sheet list.
/// </summary>
internal static class XlsxAppPropertiesPart
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace Vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    public static byte[] Generate(XlsxAppProperties props, XlsxWorkbook workbook)
    {
        // HeadingPairs: "Worksheets" → sheet count
        var headingPairs = new XElement(Vt + "vector",
            new XAttribute("size",     "2"),
            new XAttribute("baseType", "variant"),
            new XElement(Vt + "variant", new XElement(Vt + "lpstr", "Worksheets")),
            new XElement(Vt + "variant", new XElement(Vt + "i4",    workbook.Sheets.Count)));

        // TitlesOfParts: one entry per sheet name
        var titlesParts = new XElement(Vt + "vector",
            new XAttribute("size",     workbook.Sheets.Count),
            new XAttribute("baseType", "lpstr"),
            workbook.Sheets.Select(s => new XElement(Vt + "lpstr", s.Name)));

        var root = new XElement(Ns + "Properties",
            new XAttribute(XNamespace.Xmlns + "vt", Vt.NamespaceName),
            new XElement(Ns + "Application",       props.Application),
            new XElement(Ns + "DocSecurity",       props.DocSecurity),
            new XElement(Ns + "ScaleCrop",         props.ScaleCrop.ToString().ToLowerInvariant()),
            new XElement(Ns + "HeadingPairs",      headingPairs),
            new XElement(Ns + "TitlesOfParts",     titlesParts),
            new XElement(Ns + "Company",           props.Company),
            new XElement(Ns + "LinksUpToDate",     props.LinksUpToDate.ToString().ToLowerInvariant()),
            new XElement(Ns + "SharedDoc",         props.SharedDoc.ToString().ToLowerInvariant()),
            new XElement(Ns + "HyperlinksChanged", props.HyperlinksChanged.ToString().ToLowerInvariant()),
            new XElement(Ns + "AppVersion",        props.AppVersion));

        return XlsxXmlHelper.ToXmlBytes(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }
}
