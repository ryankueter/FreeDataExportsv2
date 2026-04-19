namespace FreeDataExportsv2
{
    // ── Public types ───────────────────────────────────────────────────────────────

    /// <summary>The chart display type.</summary>
    public enum ChartType
    {
        /// <summary>Vertical bar (column) chart — bars grow upward.</summary>
        Column,
        /// <summary>Horizontal bar chart — bars grow rightward.</summary>
        Bar,
        /// <summary>Line chart.</summary>
        Line,
        /// <summary>Pie chart.</summary>
        Pie,
        /// <summary>Area chart.</summary>
        Area,
    }

    /// <summary>One data series within a <see cref="ChartDefinition"/>.</summary>
    public sealed class ChartSeries
    {
        /// <summary>
        /// Series name displayed in the legend.
        /// May be a plain string (<c>"Revenue"</c>) or an A1-style sheet reference
        /// (<c>"Orders!$D$1"</c>).
        /// </summary>
        public string  Name        { get; set; } = string.Empty;

        /// <summary>
        /// A1-style reference for category (X-axis) labels,
        /// e.g. <c>"Orders!$B$2:$B$6"</c>.  May be <c>null</c> for unlabelled series.
        /// </summary>
        public string? CategoryRef { get; set; }

        /// <summary>
        /// A1-style reference for the numeric values to plot,
        /// e.g. <c>"Orders!$D$2:$D$6"</c>.
        /// </summary>
        public string  ValuesRef   { get; set; } = string.Empty;

        public ChartSeries() { }

        public ChartSeries(string name, string valuesRef, string? categoryRef = null)
        {
            Name        = name;
            ValuesRef   = valuesRef;
            CategoryRef = categoryRef;
        }
    }

    /// <summary>
    /// Describes a chart to embed in a worksheet via
    /// <see cref="XlsxWorksheet.AddChart(ChartDefinition, ObjectAnchor?)"/>.
    /// Supported by both <see cref="XlsxFile"/> and <see cref="OdsFile"/>.
    /// </summary>
    public sealed class ChartDefinition
    {
        /// <summary>Chart title displayed above the plot area. <c>null</c> = no title.</summary>
        public string?    Title          { get; set; }

        /// <summary>Chart display type (default: <see cref="ChartType.Column"/>).</summary>
        public ChartType  ChartType      { get; set; } = ChartType.Column;

        /// <summary>
        /// Legend position: <c>"r"</c> (right), <c>"l"</c> (left), <c>"t"</c> (top),
        /// <c>"b"</c> (bottom), <c>"tr"</c> (top-right).  <c>null</c> hides the legend.
        /// </summary>
        public string?    LegendPosition { get; set; } = "r";

        /// <summary>The data series to plot.</summary>
        public List<ChartSeries> DataSeries { get; } = [];

        public ChartDefinition() { }
        public ChartDefinition(string? title) { Title = title; }

        // ── Fluent builders ───────────────────────────────────────────────────────

        /// <summary>Sets the chart type.</summary>
        public ChartDefinition Type(ChartType type)     { ChartType      = type;     return this; }

        /// <summary>Sets the legend position (<c>"r"</c>, <c>"l"</c>, <c>"t"</c>, <c>"b"</c>, <c>"tr"</c>).</summary>
        public ChartDefinition Legend(string? position) { LegendPosition = position; return this; }

        /// <summary>Hides the chart legend.</summary>
        public ChartDefinition HideLegend()             { LegendPosition = null;     return this; }

        /// <summary>Appends a data series to the chart.</summary>
        /// <param name="name">Series label — plain string or sheet reference like <c>Sheet1!$B$1</c>.</param>
        /// <param name="valuesRef">Sheet reference for numeric values, e.g. <c>Sheet1!$B$2:$B$6</c>.</param>
        /// <param name="categoryRef">Optional sheet reference for category labels.</param>
        public ChartDefinition Series(string name, string valuesRef, string? categoryRef = null)
        {
            DataSeries.Add(new ChartSeries(name, valuesRef, categoryRef));
            return this;
        }
    }

    /// <summary>
    /// Two-cell anchor that controls where a chart or image is placed on the worksheet.
    /// Supported by both <see cref="XlsxFile"/> and <see cref="OdsFile"/>.
    /// </summary>
    /// <remarks>
    /// Row and column values are <b>zero-based</b>: column 0 = column A, row 0 = row 1.
    /// Offset values are in EMUs (English Metric Units); use 0 for flush alignment.
    /// </remarks>
    public sealed class ObjectAnchor
    {
        /// <summary>0-based column index for the top-left corner (default 0 = column A).</summary>
        public int FromCol    { get; set; } = 0;
        /// <summary>EMU offset from the left edge of <see cref="FromCol"/> (default 0).</summary>
        public int FromColOff { get; set; } = 0;
        /// <summary>0-based row index for the top-left corner (default 1 = row 2).</summary>
        public int FromRow    { get; set; } = 1;
        /// <summary>EMU offset from the top edge of <see cref="FromRow"/> (default 0).</summary>
        public int FromRowOff { get; set; } = 0;
        /// <summary>0-based column index for the bottom-right corner (default 7 = column H).</summary>
        public int ToCol      { get; set; } = 7;
        /// <summary>EMU offset from the left edge of <see cref="ToCol"/> (default 0).</summary>
        public int ToColOff   { get; set; } = 0;
        /// <summary>0-based row index for the bottom-right corner (default 20 = row 21).</summary>
        public int ToRow      { get; set; } = 20;
        /// <summary>EMU offset from the top edge of <see cref="ToRow"/> (default 0).</summary>
        public int ToRowOff   { get; set; } = 0;
    }
}

namespace FreeDataExportsv2.Internal
{
    /// <summary>
    /// Internal binding of a <see cref="FreeDataExportsv2.ChartDefinition"/> to a position on a sheet.
    /// IDs are assigned by <see cref="FreeDataExportsv2.XlsxFile"/> before writing.
    /// </summary>
    internal sealed class ChartInfo
    {
        public ChartDefinition Definition { get; }
        public ObjectAnchor     Anchor     { get; }

        /// <summary>Global 1-based chart ID (unique across all charts in the workbook).</summary>
        public int ChartId { get; set; }

        /// <summary>Global 1-based drawing ID for this chart's parent worksheet.</summary>
        public int DrawingId { get; set; }

        public ChartInfo(ChartDefinition definition, ObjectAnchor anchor)
        {
            Definition = definition;
            Anchor     = anchor;
        }
    }
}
