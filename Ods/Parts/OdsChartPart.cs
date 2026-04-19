using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace FreeDataExportsv2.Internal;

/// <summary>Generates Object N/content.xml, styles.xml, and meta.xml for ODS chart objects.</summary>
internal static class OdsChartPart
{
    // ── Entry points ───────────────────────────────────────────────────────────

    public static byte[] GenerateContent(ChartDefinition def, string defaultSheet)
        => Encoding.UTF8.GetBytes(BuildContent(def, defaultSheet));

    public static byte[] GenerateStyles() => Encoding.UTF8.GetBytes("""
        <?xml version="1.0" encoding="UTF-8"?>
        <office:document-styles
          xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
          office:version="1.4">
          <office:styles/>
          <office:automatic-styles/>
          <office:master-styles/>
        </office:document-styles>
        """);

    public static byte[] GenerateMeta() => Encoding.UTF8.GetBytes("""
        <?xml version="1.0" encoding="UTF-8"?>
        <office:document-meta
          xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
          office:version="1.4">
          <office:meta>
            <meta:generator>FreeDataExportsv2</meta:generator>
          </office:meta>
        </office:document-meta>
        """);

    // ── Chart content.xml ─────────────────────────────────────────────────────

    private static string BuildContent(ChartDefinition def, string defaultSheet)
    {
        string chartClass = def.ChartType switch
        {
            ChartType.Bar  => "chart:bar",
            ChartType.Line => "chart:line",
            ChartType.Pie  => "chart:circle",
            ChartType.Area => "chart:area",
            _              => "chart:bar",   // Column = bar with vertical=true
        };

        bool isVertical = def.ChartType is ChartType.Column or ChartType.Area;
        bool isPie      = def.ChartType == ChartType.Pie;

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
        sb.Append("<office:document-content\n");
        sb.Append("  xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"\n");
        sb.Append("  xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\"\n");
        sb.Append("  xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"\n");
        sb.Append("  xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\"\n");
        sb.Append("  xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\"\n");
        sb.Append("  xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"\n");
        sb.Append("  xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\"\n");
        sb.Append("  xmlns:xlink=\"http://www.w3.org/1999/xlink\"\n");
        sb.Append("  office:version=\"1.4\">\n");
        sb.Append("  <office:body>\n");
        sb.Append("    <office:chart>\n");

        // chart:chart element
        sb.Append($"      <chart:chart chart:class=\"{chartClass}\"");
        sb.Append(" svg:width=\"10cm\" svg:height=\"7cm\"");
        if (isVertical) sb.Append(" chart:column-mapping=\"\"");
        sb.Append(">\n");

        // Title
        if (!string.IsNullOrEmpty(def.Title))
        {
            sb.Append("        <chart:title><text:p>");
            sb.Append(XmlEsc(def.Title));
            sb.Append("</text:p></chart:title>\n");
        }
        else
        {
            sb.Append("        <chart:title chart:style-name=\"\"><text:p/></chart:title>\n");
        }

        // Legend
        if (def.LegendPosition != null && def.LegendPosition != "none")
        {
            string legendPos = def.LegendPosition switch
            {
                "b"  => "bottom",
                "t"  => "top",
                "l"  => "left",
                "r"  or "tr" => "right",
                _    => "right",
            };
            sb.Append($"        <chart:legend chart:legend-position=\"{legendPos}\"/>\n");
        }

        // Plot area
        sb.Append("        <chart:plot-area");
        if (isVertical) sb.Append(" chart:vertical=\"true\"");
        sb.Append(">\n");

        if (!isPie)
        {
            // Category axis (X)
            string? catRef = def.DataSeries.FirstOrDefault()?.CategoryRef;
            sb.Append("          <chart:axis chart:dimension=\"x\" chart:name=\"primary-x\">\n");
            if (!string.IsNullOrEmpty(catRef))
            {
                string odsRef = ConvertRef(catRef, defaultSheet);
                sb.Append($"            <chart:categories table:cell-range-address=\"{XmlEsc(odsRef)}\"/>\n");
            }
            sb.Append("          </chart:axis>\n");

            // Value axis (Y)
            sb.Append("          <chart:axis chart:dimension=\"y\" chart:name=\"primary-y\"/>\n");
        }

        // Series
        foreach (var series in def.DataSeries)
        {
            string valRef = ConvertRef(series.ValuesRef, defaultSheet);
            sb.Append($"          <chart:series chart:class=\"{chartClass}\"");
            sb.Append($" chart:values-cell-range-address=\"{XmlEsc(valRef)}\"");

            if (!string.IsNullOrEmpty(series.Name))
            {
                // Try to interpret as a cell reference first, otherwise treat as literal
                string nameRef = series.Name.Contains('!') || series.Name.Contains('$')
                    ? ConvertRef(series.Name, defaultSheet)
                    : "";
                if (!string.IsNullOrEmpty(nameRef))
                    sb.Append($" chart:label-cell-address=\"{XmlEsc(nameRef)}\"");
            }
            sb.Append(">\n");

            // Data points — count from valuesRef range if possible
            int pointCount = CountRangePoints(series.ValuesRef);
            if (pointCount > 0)
                sb.Append($"            <chart:data-point chart:repeated=\"{pointCount}\"/>\n");

            sb.Append("          </chart:series>\n");
        }

        sb.Append("        </chart:plot-area>\n");
        sb.Append("      </chart:chart>\n");
        sb.Append("    </office:chart>\n");
        sb.Append("  </office:body>\n");
        sb.Append("</office:document-content>\n");
        return sb.ToString();
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Converts an XLSX-style cell reference (e.g. "Sheet1!$A$2:$A$6") to
    /// ODS chart format (e.g. "$Sheet1.$A$2:$Sheet1.$A$6").
    /// </summary>
    internal static string ConvertRef(string xlsxRef, string defaultSheet)
    {
        if (string.IsNullOrWhiteSpace(xlsxRef)) return xlsxRef;

        // Split on '!'
        string sheet;
        string range;
        int bang = xlsxRef.IndexOf('!');
        if (bang >= 0)
        {
            sheet = xlsxRef[..bang].Trim('$', '\'', '"');
            range = xlsxRef[(bang + 1)..];
        }
        else
        {
            sheet = defaultSheet;
            range = xlsxRef;
        }

        // Split on ':'
        int colon = range.IndexOf(':');
        if (colon < 0)
            return $"${sheet}.{range}";

        string from = range[..colon];
        string to   = range[(colon + 1)..];
        // If the 'to' part has no sheet prefix, repeat the sheet name
        return $"${sheet}.{from}:${sheet}.{to}";
    }

    /// <summary>Counts data points in a range like "Sheet1!$A$2:$A$6" → 5.</summary>
    private static int CountRangePoints(string valuesRef)
    {
        try
        {
            int bang  = valuesRef.IndexOf('!');
            string range = bang >= 0 ? valuesRef[(bang + 1)..] : valuesRef;
            int colon = range.IndexOf(':');
            if (colon < 0) return 1;
            string from = range[..colon].TrimStart('$');
            string to   = range[(colon + 1)..].TrimStart('$');

            // Extract row numbers
            var fromRow = int.Parse(Regex.Match(from, @"\d+").Value);
            var toRow   = int.Parse(Regex.Match(to,   @"\d+").Value);
            return Math.Abs(toRow - fromRow) + 1;
        }
        catch { return 0; }
    }

    private static string XmlEsc(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
}
