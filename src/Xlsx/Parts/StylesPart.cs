using System.Xml.Linq;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates xl/styles.xml from a <see cref="XlsxStyles"/>.
/// </summary>
internal static class XlsxStylesPart
{
    private static readonly XNamespace Ns    = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Mc    = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace X14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private static readonly XNamespace X14   = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private static readonly XNamespace X15   = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
    private static readonly XNamespace X16r2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main";
    private static readonly XNamespace Xr    = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

    public static byte[] Generate(XlsxStyles styles)
    {
        var children = new List<object?>();

        // <numFmts> — only written when custom formats exist
        if (styles.NumFmts.Count > 0)
        {
            children.Add(new XElement(Ns + "numFmts",
                new XAttribute("count", styles.NumFmts.Count),
                styles.NumFmts.Select(f =>
                    new XElement(Ns + "numFmt",
                        new XAttribute("numFmtId",   f.NumFmtId),
                        new XAttribute("formatCode", f.FormatCode)))));
        }

        children.Add(new XElement(Ns + "fonts",
            new XAttribute("count", styles.Fonts.Count),
            new XAttribute(X14ac + "knownFonts", "1"),
            styles.Fonts.Select(BuildFont)));

        children.Add(new XElement(Ns + "fills",
            new XAttribute("count", styles.Fills.Count),
            styles.Fills.Select(BuildFill)));

        children.Add(new XElement(Ns + "borders",
            new XAttribute("count", styles.Borders.Count),
            styles.Borders.Select(BuildBorder)));

        children.Add(new XElement(Ns + "cellStyleXfs",
            new XAttribute("count", styles.CellStyleXfs.Count),
            styles.CellStyleXfs.Select(x => BuildXf(x))));

        children.Add(new XElement(Ns + "cellXfs",
            new XAttribute("count", styles.CellXfs.Count),
            styles.CellXfs.Select(x => BuildXf(x))));

        children.Add(new XElement(Ns + "cellStyles",
            new XAttribute("count", styles.CellStyles.Count),
            styles.CellStyles.Select(cs =>
                new XElement(Ns + "cellStyle",
                    new XAttribute("name",      cs.Name),
                    new XAttribute("xfId",      cs.XfId),
                    new XAttribute("builtinId", cs.BuiltinId)))));

        children.Add(new XElement(Ns + "dxfs",  new XAttribute("count", "0")));
        children.Add(new XElement(Ns + "tableStyles",
            new XAttribute("count",             "0"),
            new XAttribute("defaultTableStyle", "TableStyleMedium2"),
            new XAttribute("defaultPivotStyle", "PivotStyleLight16")));

        children.Add(new XElement(Ns + "extLst",
            new XElement(Ns + "ext",
                new XAttribute("uri", "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"),
                new XAttribute(XNamespace.Xmlns + "x14", X14.NamespaceName),
                new XElement(X14 + "slicerStyles",
                    new XAttribute("defaultSlicerStyle", "SlicerStyleLight1"))),
            new XElement(Ns + "ext",
                new XAttribute("uri", "{9260A510-F301-46a8-8635-F512D64BE5F5}"),
                new XAttribute(XNamespace.Xmlns + "x15", X15.NamespaceName),
                new XElement(X15 + "timelineStyles",
                    new XAttribute("defaultTimelineStyle", "TimeSlicerStyleLight1")))));

        var root = new XElement(Ns + "styleSheet",
            new XAttribute(XNamespace.Xmlns + "mc",    Mc.NamespaceName),
            new XAttribute(Mc + "Ignorable",            "x14ac x16r2 xr"),
            new XAttribute(XNamespace.Xmlns + "x14ac", X14ac.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "x16r2", X16r2.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xr",    Xr.NamespaceName),
            children);

        return XlsxXmlHelper.ToXmlBytes(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }

    // ── Element builders ──────────────────────────────────────────────────────

    private static XElement BuildFont(XlsxFont f)
    {
        var el = new XElement(Ns + "font");
        if (f.Bold)          el.Add(new XElement(Ns + "b"));
        if (f.Italic)        el.Add(new XElement(Ns + "i"));
        if (f.Strikethrough) el.Add(new XElement(Ns + "strike"));
        if (f.Underline)     el.Add(new XElement(Ns + "u"));
        el.Add(new XElement(Ns + "sz", new XAttribute("val", XlsxXmlHelper.F(f.Size))));
        if (f.ThemeColor.HasValue)
            el.Add(new XElement(Ns + "color", new XAttribute("theme", f.ThemeColor.Value)));
        else if (f.RgbColor is not null)
            el.Add(new XElement(Ns + "color", new XAttribute("rgb", f.RgbColor)));
        el.Add(new XElement(Ns + "name",   new XAttribute("val", f.Name)));
        el.Add(new XElement(Ns + "family", new XAttribute("val", f.Family)));
        if (f.Scheme is not null)
            el.Add(new XElement(Ns + "scheme", new XAttribute("val", f.Scheme)));
        return el;
    }

    private static XElement BuildFill(XlsxFill f)
    {
        var pattern = new XElement(Ns + "patternFill", new XAttribute("patternType", f.PatternType));
        if (f.FgColor is not null) pattern.Add(new XElement(Ns + "fgColor", new XAttribute("rgb", f.FgColor)));
        if (f.BgColor is not null) pattern.Add(new XElement(Ns + "bgColor", new XAttribute("rgb", f.BgColor)));
        return new XElement(Ns + "fill", pattern);
    }

    private static XElement BuildBorder(XlsxBorder b)
    {
        return new XElement(Ns + "border",
            Side("left",     b.Left),
            Side("right",    b.Right),
            Side("top",      b.Top),
            Side("bottom",   b.Bottom),
            Side("diagonal", b.Diagonal));

        XElement Side(string name, XlsxBorderSide s)
        {
            var el = new XElement(Ns + name);
            if (s.Style is not null) el.Add(new XAttribute("style", s.Style));
            if (s.Color is not null) el.Add(new XElement(Ns + "color", new XAttribute("rgb", s.Color)));
            return el;
        }
    }

    private static XElement BuildXf(XlsxCellXf x)
    {
        var el = new XElement(Ns + "xf",
            new XAttribute("numFmtId", x.NumFmtId),
            new XAttribute("fontId",   x.FontId),
            new XAttribute("fillId",   x.FillId),
            new XAttribute("borderId", x.BorderId));
        if (x.XfId.HasValue)              el.Add(new XAttribute("xfId",              x.XfId.Value));
        if (x.ApplyNumberFormat)          el.Add(new XAttribute("applyNumberFormat", "1"));
        if (x.ApplyFont)                  el.Add(new XAttribute("applyFont",         "1"));
        if (x.ApplyFill)                  el.Add(new XAttribute("applyFill",         "1"));
        if (x.ApplyBorder)                el.Add(new XAttribute("applyBorder",       "1"));
        if (x.ApplyAlignment)             el.Add(new XAttribute("applyAlignment",    "1"));
        if (x.XlsxAlignment is not null)
        {
            var a = new XElement(Ns + "alignment");
            if (x.XlsxAlignment.Horizontal is not null) a.Add(new XAttribute("horizontal", x.XlsxAlignment.Horizontal));
            if (x.XlsxAlignment.Vertical   is not null) a.Add(new XAttribute("vertical",   x.XlsxAlignment.Vertical));
            if (x.XlsxAlignment.WrapText)               a.Add(new XAttribute("wrapText",   "1"));
            el.Add(a);
        }
        return el;
    }
}
