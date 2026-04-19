using System.Globalization;
using System.Xml;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates xl/charts/chart{N}.xml — the Open XML chart definition.
/// </summary>
internal static class XlsxChartPart
{
    // Namespace constants
    private const string NsC  = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string NsA  = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string NsR  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static byte[] Generate(ChartDefinition chart)
    {
        using var ms  = new System.IO.MemoryStream();
        using var xml = XmlWriter.Create(ms, new XmlWriterSettings
        {
            Encoding           = new System.Text.UTF8Encoding(false),
            Indent             = false,
            OmitXmlDeclaration = false,
        });

        xml.WriteStartDocument(true);

        // <c:chartSpace>
        xml.WriteStartElement("c", "chartSpace", NsC);
        xml.WriteAttributeString("xmlns", "c", null, NsC);
        xml.WriteAttributeString("xmlns", "a", null, NsA);
        xml.WriteAttributeString("xmlns", "r", null, NsR);

        // <c:lang val="en-US"/>
        xml.WriteStartElement("c", "lang", NsC);
        xml.WriteAttributeString("val", "en-US");
        xml.WriteEndElement();

        // <c:roundedCorners val="0"/>
        xml.WriteStartElement("c", "roundedCorners", NsC);
        xml.WriteAttributeString("val", "0");
        xml.WriteEndElement();

        // <c:chart>
        xml.WriteStartElement("c", "chart", NsC);

        // Title
        if (chart.Title is not null)
        {
            xml.WriteStartElement("c", "title", NsC);
            xml.WriteStartElement("c", "tx", NsC);
            xml.WriteStartElement("c", "rich", NsC);
            xml.WriteStartElement("a", "bodyPr", NsA); xml.WriteEndElement();
            xml.WriteStartElement("a", "lstStyle", NsA); xml.WriteEndElement();
            xml.WriteStartElement("a", "p", NsA);
            xml.WriteStartElement("a", "r", NsA);
            xml.WriteStartElement("a", "t", NsA);
            xml.WriteString(chart.Title);
            xml.WriteEndElement(); // a:t
            xml.WriteEndElement(); // a:r
            xml.WriteEndElement(); // a:p
            xml.WriteEndElement(); // c:rich
            xml.WriteEndElement(); // c:tx
            xml.WriteStartElement("c", "overlay", NsC);
            xml.WriteAttributeString("val", "0");
            xml.WriteEndElement();
            xml.WriteEndElement(); // c:title
        }

        xml.WriteStartElement("c", "autoTitleDeleted", NsC);
        xml.WriteAttributeString("val", chart.Title is null ? "1" : "0");
        xml.WriteEndElement();

        // <c:plotArea>
        xml.WriteStartElement("c", "plotArea", NsC);

        bool hasCatValAxes = chart.ChartType != ChartType.Pie;
        WriteChartBody(xml, chart, hasCatValAxes);

        xml.WriteEndElement(); // c:plotArea

        // <c:legend>
        if (chart.LegendPosition is not null)
        {
            xml.WriteStartElement("c", "legend", NsC);
            xml.WriteStartElement("c", "legendPos", NsC);
            xml.WriteAttributeString("val", chart.LegendPosition);
            xml.WriteEndElement();
            xml.WriteStartElement("c", "overlay", NsC);
            xml.WriteAttributeString("val", "0");
            xml.WriteEndElement();
            xml.WriteEndElement(); // c:legend
        }

        xml.WriteStartElement("c", "plotVisOnly", NsC);
        xml.WriteAttributeString("val", "1");
        xml.WriteEndElement();

        xml.WriteEndElement(); // c:chart
        xml.WriteEndElement(); // c:chartSpace

        xml.Flush();
        return ms.ToArray();
    }

    // ── Chart body (chart-type element + axes) ────────────────────────────────

    private static void WriteChartBody(XmlWriter xml, ChartDefinition chart, bool hasCatValAxes)
    {
        switch (chart.ChartType)
        {
            case ChartType.Column:
                WriteBarChart(xml, chart, barDir: "col");
                break;
            case ChartType.Bar:
                WriteBarChart(xml, chart, barDir: "bar");
                break;
            case ChartType.Line:
                WriteLineChart(xml, chart);
                break;
            case ChartType.Pie:
                WritePieChart(xml, chart);
                break;
            case ChartType.Area:
                WriteAreaChart(xml, chart);
                break;
        }

        if (hasCatValAxes)
            WriteCatValAxes(xml, chart.ChartType == ChartType.Bar);
    }

    // ── Bar / Column chart ────────────────────────────────────────────────────

    private static void WriteBarChart(XmlWriter xml, ChartDefinition chart, string barDir)
    {
        xml.WriteStartElement("c", "barChart", NsC);

        xml.WriteStartElement("c", "barDir", NsC);
        xml.WriteAttributeString("val", barDir);
        xml.WriteEndElement();

        xml.WriteStartElement("c", "grouping", NsC);
        xml.WriteAttributeString("val", "clustered");
        xml.WriteEndElement();

        xml.WriteStartElement("c", "varyColors", NsC);
        xml.WriteAttributeString("val", "0");
        xml.WriteEndElement();

        WriteSeries(xml, chart.DataSeries);

        WriteAxIds(xml);
        xml.WriteEndElement(); // c:barChart
    }

    // ── Line chart ────────────────────────────────────────────────────────────

    private static void WriteLineChart(XmlWriter xml, ChartDefinition chart)
    {
        xml.WriteStartElement("c", "lineChart", NsC);

        xml.WriteStartElement("c", "grouping", NsC);
        xml.WriteAttributeString("val", "standard");
        xml.WriteEndElement();

        xml.WriteStartElement("c", "varyColors", NsC);
        xml.WriteAttributeString("val", "0");
        xml.WriteEndElement();

        WriteSeries(xml, chart.DataSeries);

        xml.WriteStartElement("c", "smooth", NsC);
        xml.WriteAttributeString("val", "0");
        xml.WriteEndElement();

        WriteAxIds(xml);
        xml.WriteEndElement(); // c:lineChart
    }

    // ── Pie chart ─────────────────────────────────────────────────────────────

    private static void WritePieChart(XmlWriter xml, ChartDefinition chart)
    {
        xml.WriteStartElement("c", "pieChart", NsC);

        xml.WriteStartElement("c", "varyColors", NsC);
        xml.WriteAttributeString("val", "1");
        xml.WriteEndElement();

        WriteSeries(xml, chart.DataSeries);

        xml.WriteStartElement("c", "firstSliceAng", NsC);
        xml.WriteAttributeString("val", "0");
        xml.WriteEndElement();

        xml.WriteEndElement(); // c:pieChart
    }

    // ── Area chart ────────────────────────────────────────────────────────────

    private static void WriteAreaChart(XmlWriter xml, ChartDefinition chart)
    {
        xml.WriteStartElement("c", "areaChart", NsC);

        xml.WriteStartElement("c", "grouping", NsC);
        xml.WriteAttributeString("val", "standard");
        xml.WriteEndElement();

        xml.WriteStartElement("c", "varyColors", NsC);
        xml.WriteAttributeString("val", "0");
        xml.WriteEndElement();

        WriteSeries(xml, chart.DataSeries);

        WriteAxIds(xml);
        xml.WriteEndElement(); // c:areaChart
    }

    // ── Shared helpers ────────────────────────────────────────────────────────

    private static void WriteAxIds(XmlWriter xml)
    {
        xml.WriteStartElement("c", "axId", NsC); xml.WriteAttributeString("val", "1"); xml.WriteEndElement();
        xml.WriteStartElement("c", "axId", NsC); xml.WriteAttributeString("val", "2"); xml.WriteEndElement();
    }

    private static void WriteSeries(XmlWriter xml, IReadOnlyList<ChartSeries> series)
    {
        for (int i = 0; i < series.Count; i++)
        {
            var s = series[i];
            xml.WriteStartElement("c", "ser", NsC);

            // idx / order
            xml.WriteStartElement("c", "idx", NsC); xml.WriteAttributeString("val", I(i)); xml.WriteEndElement();
            xml.WriteStartElement("c", "order", NsC); xml.WriteAttributeString("val", I(i)); xml.WriteEndElement();

            // Series name
            if (!string.IsNullOrEmpty(s.Name))
            {
                xml.WriteStartElement("c", "tx", NsC);
                if (IsSheetRef(s.Name))
                {
                    xml.WriteStartElement("c", "strRef", NsC);
                    xml.WriteStartElement("c", "f", NsC); xml.WriteString(s.Name); xml.WriteEndElement();
                    xml.WriteEndElement(); // strRef
                }
                else
                {
                    xml.WriteStartElement("c", "v", NsC); xml.WriteString(s.Name); xml.WriteEndElement();
                }
                xml.WriteEndElement(); // tx
            }

            // Category labels
            if (s.CategoryRef is not null)
            {
                xml.WriteStartElement("c", "cat", NsC);
                xml.WriteStartElement("c", "strRef", NsC);
                xml.WriteStartElement("c", "f", NsC); xml.WriteString(s.CategoryRef); xml.WriteEndElement();
                xml.WriteEndElement(); // strRef
                xml.WriteEndElement(); // cat
            }

            // Values
            if (!string.IsNullOrEmpty(s.ValuesRef))
            {
                xml.WriteStartElement("c", "val", NsC);
                xml.WriteStartElement("c", "numRef", NsC);
                xml.WriteStartElement("c", "f", NsC); xml.WriteString(s.ValuesRef); xml.WriteEndElement();
                xml.WriteEndElement(); // numRef
                xml.WriteEndElement(); // val
            }

            xml.WriteEndElement(); // ser
        }
    }

    /// <summary>Write category axis (axId=1) and value axis (axId=2).</summary>
    private static void WriteCatValAxes(XmlWriter xml, bool isHorizontalBar)
    {
        // Category axis
        xml.WriteStartElement("c", "catAx", NsC);
        xml.WriteStartElement("c", "axId", NsC);  xml.WriteAttributeString("val", "1"); xml.WriteEndElement();
        xml.WriteStartElement("c", "scaling", NsC);
        xml.WriteStartElement("c", "orientation", NsC); xml.WriteAttributeString("val", "minMax"); xml.WriteEndElement();
        xml.WriteEndElement(); // scaling
        xml.WriteStartElement("c", "delete", NsC); xml.WriteAttributeString("val", "0"); xml.WriteEndElement();
        xml.WriteStartElement("c", "axPos", NsC);
        xml.WriteAttributeString("val", isHorizontalBar ? "l" : "b");
        xml.WriteEndElement();
        xml.WriteStartElement("c", "crossAx", NsC); xml.WriteAttributeString("val", "2"); xml.WriteEndElement();
        xml.WriteEndElement(); // catAx

        // Value axis
        xml.WriteStartElement("c", "valAx", NsC);
        xml.WriteStartElement("c", "axId", NsC);  xml.WriteAttributeString("val", "2"); xml.WriteEndElement();
        xml.WriteStartElement("c", "scaling", NsC);
        xml.WriteStartElement("c", "orientation", NsC); xml.WriteAttributeString("val", "minMax"); xml.WriteEndElement();
        xml.WriteEndElement(); // scaling
        xml.WriteStartElement("c", "delete", NsC); xml.WriteAttributeString("val", "0"); xml.WriteEndElement();
        xml.WriteStartElement("c", "axPos", NsC);
        xml.WriteAttributeString("val", isHorizontalBar ? "b" : "l");
        xml.WriteEndElement();
        xml.WriteStartElement("c", "numFmt", NsC);
        xml.WriteAttributeString("formatCode", "General");
        xml.WriteAttributeString("sourceLinked", "0");
        xml.WriteEndElement();
        xml.WriteStartElement("c", "tickLblPos", NsC); xml.WriteAttributeString("val", "nextTo"); xml.WriteEndElement();
        xml.WriteStartElement("c", "crossAx", NsC); xml.WriteAttributeString("val", "1"); xml.WriteEndElement();
        xml.WriteEndElement(); // valAx
    }

    private static bool IsSheetRef(string s) => s.Contains('!');
    private static string I(int v) => v.ToString(CultureInfo.InvariantCulture);
}
