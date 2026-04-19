namespace FreeDataExportsv2
{
    // ── Public types ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Defines an Excel table to be attached to a worksheet via
    /// <see cref="XlsxWorksheet.AddTable(string, XlsxTableDefinition)"/>.
    /// </summary>
    public sealed class XlsxTableDefinition
    {
        /// <summary>Internal table name (used in structured references, e.g. Table1[Column]).</summary>
        public string Name           { get; set; }
        /// <summary>Display name shown in Excel's Name Manager. Defaults to <see cref="Name"/>.</summary>
        public string DisplayName    { get; set; }
        /// <summary>Table style. Use <see cref="XlsxTableStyles"/> constants or any valid Excel style name.</summary>
        public string StyleName      { get; set; } = XlsxTableStyles.Medium2;
        /// <summary>When true, adds a totals row at the bottom of the table range.</summary>
        public bool   HasTotalsRow   { get; set; }
        public bool   ShowFirstColumn{ get; set; }
        public bool   ShowLastColumn { get; set; }
        /// <summary>Alternating row shading (default: true).</summary>
        public bool   ShowRowStripes { get; set; } = true;
        public bool   ShowColStripes { get; set; }

        public List<XlsxTableColumn> Columns { get; } = [];

        public XlsxTableDefinition(string name)
        {
            Name        = name;
            DisplayName = name;
        }

        // ── Fluent builders ───────────────────────────────────────────────────────

        /// <summary>Appends a column definition.</summary>
        public XlsxTableDefinition AddColumn(string name,
            XlsxTotalsRowFunction totalsFunction = XlsxTotalsRowFunction.None)
        {
            Columns.Add(new XlsxTableColumn(name, totalsFunction));
            return this;
        }

        public XlsxTableDefinition Style(string styleName)            { StyleName      = styleName; return this; }
        public XlsxTableDefinition ShowTotalsRow(bool has = true)    { HasTotalsRow   = has;       return this; }
        public XlsxTableDefinition RowStripes(bool show = true)      { ShowRowStripes = show;      return this; }
        public XlsxTableDefinition ColStripes(bool show = true)      { ShowColStripes = show;      return this; }
        public XlsxTableDefinition FirstColumn(bool show = true)     { ShowFirstColumn = show;     return this; }
        public XlsxTableDefinition LastColumn(bool show = true)      { ShowLastColumn  = show;     return this; }
    }

    /// <summary>One column in a <see cref="XlsxTableDefinition"/>.</summary>
    public sealed class XlsxTableColumn
    {
        public string            Name           { get; set; }
        public XlsxTotalsRowFunction TotalsFunction { get; set; }

        public XlsxTableColumn(string name, XlsxTotalsRowFunction totalsFunction = XlsxTotalsRowFunction.None)
        {
            Name           = name;
            TotalsFunction = totalsFunction;
        }
    }

    /// <summary>Aggregate function displayed in the table's totals row.</summary>
    public enum XlsxTotalsRowFunction
    {
        None,
        Sum,
        Average,
        Count,
        CountNums,
        Max,
        Min,
        StdDev,
        Var,
    }

    /// <summary>
    /// Named constants for every built-in Excel table style.
    /// Pass any value to <see cref="XlsxTableDefinition.StyleName"/> or <see cref="XlsxTableDefinition.Style"/>.
    /// </summary>
    public static class XlsxTableStyles
    {
        // ── Light ─────────────────────────────────────────────────────────────────
        public const string Light1  = "TableStyleLight1";
        public const string Light2  = "TableStyleLight2";
        public const string Light3  = "TableStyleLight3";
        public const string Light4  = "TableStyleLight4";
        public const string Light5  = "TableStyleLight5";
        public const string Light6  = "TableStyleLight6";
        public const string Light7  = "TableStyleLight7";
        public const string Light8  = "TableStyleLight8";
        public const string Light9  = "TableStyleLight9";
        public const string Light10 = "TableStyleLight10";
        public const string Light11 = "TableStyleLight11";
        public const string Light12 = "TableStyleLight12";
        public const string Light13 = "TableStyleLight13";
        public const string Light14 = "TableStyleLight14";
        public const string Light15 = "TableStyleLight15";
        public const string Light16 = "TableStyleLight16";
        public const string Light17 = "TableStyleLight17";
        public const string Light18 = "TableStyleLight18";
        public const string Light19 = "TableStyleLight19";
        public const string Light20 = "TableStyleLight20";
        public const string Light21 = "TableStyleLight21";

        // ── Medium ────────────────────────────────────────────────────────────────
        public const string Medium1  = "TableStyleMedium1";
        public const string Medium2  = "TableStyleMedium2";
        public const string Medium3  = "TableStyleMedium3";
        public const string Medium4  = "TableStyleMedium4";
        public const string Medium5  = "TableStyleMedium5";
        public const string Medium6  = "TableStyleMedium6";
        public const string Medium7  = "TableStyleMedium7";
        public const string Medium8  = "TableStyleMedium8";
        public const string Medium9  = "TableStyleMedium9";
        public const string Medium10 = "TableStyleMedium10";
        public const string Medium11 = "TableStyleMedium11";
        public const string Medium12 = "TableStyleMedium12";
        public const string Medium13 = "TableStyleMedium13";
        public const string Medium14 = "TableStyleMedium14";
        public const string Medium15 = "TableStyleMedium15";
        public const string Medium16 = "TableStyleMedium16";
        public const string Medium17 = "TableStyleMedium17";
        public const string Medium18 = "TableStyleMedium18";
        public const string Medium19 = "TableStyleMedium19";
        public const string Medium20 = "TableStyleMedium20";
        public const string Medium21 = "TableStyleMedium21";
        public const string Medium22 = "TableStyleMedium22";
        public const string Medium23 = "TableStyleMedium23";
        public const string Medium24 = "TableStyleMedium24";
        public const string Medium25 = "TableStyleMedium25";
        public const string Medium26 = "TableStyleMedium26";
        public const string Medium27 = "TableStyleMedium27";
        public const string Medium28 = "TableStyleMedium28";

        // ── Dark ──────────────────────────────────────────────────────────────────
        public const string Dark1  = "TableStyleDark1";
        public const string Dark2  = "TableStyleDark2";
        public const string Dark3  = "TableStyleDark3";
        public const string Dark4  = "TableStyleDark4";
        public const string Dark5  = "TableStyleDark5";
        public const string Dark6  = "TableStyleDark6";
        public const string Dark7  = "TableStyleDark7";
        public const string Dark8  = "TableStyleDark8";
        public const string Dark9  = "TableStyleDark9";
        public const string Dark10 = "TableStyleDark10";
        public const string Dark11 = "TableStyleDark11";
    }
}

namespace FreeDataExportsv2.Internal
{
    /// <summary>
    /// Internal binding of a <see cref="FreeDataExportsv2.XlsxTableDefinition"/> to a cell range within a sheet.
    /// Global table ID and per-sheet relationship ID are assigned by <see cref="FreeDataExportsv2.XlsxFile"/>
    /// immediately before writing the ZIP.
    /// </summary>
    internal sealed class XlsxTableInfo
    {
        public string          Range      { get; }
        public XlsxTableDefinition Definition { get; }

        /// <summary>Global 1-based table ID (unique across all tables in the workbook).</summary>
        public int TableId  { get; set; }

        /// <summary>1-based relationship ID within this sheet's _rels file (rId1, rId2, …).</summary>
        public int LocalRId { get; set; }

        public XlsxTableInfo(string range, XlsxTableDefinition definition)
        {
            Range      = range;
            Definition = definition;
        }
    }
}
