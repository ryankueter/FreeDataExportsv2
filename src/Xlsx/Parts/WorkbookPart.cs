using System.Globalization;
using System.Xml.Linq;

namespace FreeDataExportsv2.Internal;

/// <summary>Generates xl/workbook.xml.</summary>
internal static class XlsxWorkbookPart
{
    private static readonly XNamespace Ns     = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace NsR    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace NsMc   = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace NsX15  = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
    private static readonly XNamespace NsXr   = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private static readonly XNamespace NsXr2  = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
    private static readonly XNamespace NsXr6  = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6";
    private static readonly XNamespace NsXr10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";

    public static byte[] Generate(XlsxWorkbook workbook)
    {
        var sheets = workbook.Sheets.Select((s, i) =>
            new XElement(Ns + "sheet",
                new XAttribute("name",    s.Name),
                new XAttribute("sheetId", (i + 1).ToString(CultureInfo.InvariantCulture)),
                new XAttribute(NsR + "id", $"rId{i + 1}")));

        var root = new XElement(Ns + "workbook",
            new XAttribute(XNamespace.Xmlns + "r",    NsR.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "mc",   NsMc.NamespaceName),
            new XAttribute(NsMc + "Ignorable",        "x15 xr xr6 xr10 xr2"),
            new XAttribute(XNamespace.Xmlns + "x15",  NsX15.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xr",   NsXr.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xr2",  NsXr2.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xr6",  NsXr6.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xr10", NsXr10.NamespaceName),
            new XElement(Ns + "fileVersion",
                new XAttribute("appName",      "xl"),
                new XAttribute("lastEdited",   "7"),
                new XAttribute("lowestEdited", "7"),
                new XAttribute("rupBuild",     "22228")),
            new XElement(Ns + "workbookPr",
                new XAttribute("defaultThemeVersion", "166925")),
            new XElement(NsMc + "AlternateContent",
                new XAttribute(XNamespace.Xmlns + "mc", NsMc.NamespaceName),
                new XElement(NsMc + "Choice",
                    new XAttribute("Requires", "x15"),
                    new XElement(NsX15 + "absPath",
                        new XAttribute("url", "/")))),
            new XElement(Ns + "bookViews",
                new XElement(Ns + "workbookView",
                    new XAttribute("xWindow",       "-120"),
                    new XAttribute("yWindow",       "-120"),
                    new XAttribute("windowWidth",   "29040"),
                    new XAttribute("windowHeight",  "15840"),
                    new XAttribute(NsXr + "uid", "{00000000-0001-0000-0000-000000000000}"))),
            new XElement(Ns + "sheets", sheets),
            new XElement(Ns + "calcPr",
                new XAttribute("calcId", "191029")));

        return XlsxXmlHelper.ToXmlBytes(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }
}
