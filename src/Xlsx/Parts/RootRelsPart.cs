using System.Xml.Linq;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates _rels/.rels — the root package relationships file.
/// Links the package root to the workbook, core properties, and app properties.
/// </summary>
internal static class XlsxRootRelsPart
{
    private static readonly XNamespace Ns =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string OfficeDocRels =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageMeta =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata";

    public static byte[] Generate()
    {
        // IDs and order match the original Book1.xlsx exactly
        var root = new XElement(Ns + "Relationships",
            Rel("rId3", $"{OfficeDocRels}/extended-properties",      "docProps/app.xml"),
            Rel("rId2", $"{PackageMeta}/core-properties",             "docProps/core.xml"),
            Rel("rId1", $"{OfficeDocRels}/officeDocument",            "xl/workbook.xml"));

        return XlsxXmlHelper.ToXmlBytes(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }

    private static XElement Rel(string id, string type, string target) =>
        new(Ns + "Relationship",
            new XAttribute("Id",     id),
            new XAttribute("Type",   type),
            new XAttribute("Target", target));
}
