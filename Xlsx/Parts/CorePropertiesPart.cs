using System.Xml.Linq;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates docProps/core.xml — OPC / Dublin Core core properties.
/// </summary>
internal static class XlsxCorePropertiesPart
{
    private static readonly XNamespace Cp      = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace Dc      = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace Dcterms = "http://purl.org/dc/terms/";
    private static readonly XNamespace Dcmitype = "http://purl.org/dc/dcmitype/";
    private static readonly XNamespace Xsi     = "http://www.w3.org/2001/XMLSchema-instance";

    public static byte[] Generate(XlsxCoreProperties props)
    {
        var children = new List<object>();

        if (props.Title       != null) children.Add(new XElement(Dc + "title",       props.Title));
        if (props.Subject     != null) children.Add(new XElement(Dc + "subject",     props.Subject));
        if (props.Description != null) children.Add(new XElement(Dc + "description", props.Description));
        if (props.Keywords    != null) children.Add(new XElement(Cp + "keywords",    props.Keywords));

        children.Add(new XElement(Dc + "creator", props.Creator));
        children.Add(new XElement(Cp + "lastModifiedBy", props.LastModifiedBy));
        children.Add(new XElement(Dcterms + "created",
            new XAttribute(Xsi + "type", "dcterms:W3CDTF"),
            props.Created.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")));
        children.Add(new XElement(Dcterms + "modified",
            new XAttribute(Xsi + "type", "dcterms:W3CDTF"),
            props.Modified.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")));

        var root = new XElement(Cp + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp",      Cp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dc",      Dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcterms", Dcterms.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcmitype", Dcmitype.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xsi",     Xsi.NamespaceName),
            children);

        return XlsxXmlHelper.ToXmlBytes(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }
}
