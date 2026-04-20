using System.Globalization;
using System.Text;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates content.xml for an ODS package.
/// Translates the shared <see cref="XlsxWorkbook"/> / <see cref="XlsxStyles"/> data
/// (which also backs <see cref="XlsxFile"/>) into OpenDocument XML.
/// </summary>
internal static class OdsContentPart
{
    // ── Column/row geometry defaults (Excel character units → inches) ──────────
    private const double ColWidthPerChar = 0.0889; // inches per Excel char-width unit
    private const double DefaultRowH     = 0.2409; // inches  (≈15pt row)

    // ── Public entry point ─────────────────────────────────────────────────────

    public static byte[] Generate(XlsxWorkbook workbook,
                                   Dictionary<DataType, string> formatOverrides)
    {
        var styles   = workbook.XlsxStyles;
        var numFmtLookup = BuildNumFmtLookup(styles);

        // ── 1. Collect unique XLSX style indices used across all cells ─────────
        var usedXfIndices = new SortedSet<int>();
        foreach (var sheet in workbook.Sheets)
            foreach (var row in sheet.Rows)
                foreach (var cell in row.Cells)
                    if (cell.StyleIndex.HasValue)
                        usedXfIndices.Add(cell.StyleIndex.Value);

        // ── 2. Collect unique column widths ───────────────────────────────────
        // Map width (in Excel chars) → ODS style name
        var colWidthStyles = new Dictionary<double, string>();
        foreach (var sheet in workbook.Sheets)
            foreach (var (_, _, w) in sheet.ColWidths)
                if (!colWidthStyles.ContainsKey(w))
                    colWidthStyles[w] = $"co{colWidthStyles.Count + 1}";
        // Always add a default column style
        if (!colWidthStyles.ContainsKey(-1))
            colWidthStyles[-1] = "co0"; // sentinel for "default"

        // ── 3. Build content ──────────────────────────────────────────────────
        var sb = new StringBuilder(256 * 1024);

        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
        AppendDocumentOpen(sb);

        // font faces
        AppendFontFaces(sb, styles);

        // automatic styles: number formats, cell styles, column/row/table/graphic styles
        sb.Append("  <office:automatic-styles>\n");
        AppendNumberFormats(sb, usedXfIndices, styles, numFmtLookup);
        AppendCellStyles(sb, usedXfIndices, styles, numFmtLookup);
        AppendColumnStyles(sb, colWidthStyles);
        AppendSheetDefaultStyles(sb, workbook.Sheets);
        sb.Append("    <style:style style:name=\"ro1\" style:family=\"table-row\">\n");
        sb.Append($"      <style:table-row-properties style:row-height=\"{DefaultRowH:F4}in\" fo:break-before=\"auto\" style:use-optimal-row-height=\"true\"/>\n");
        sb.Append("    </style:style>\n");
        // Table styles for each sheet (tab colors)
        for (int i = 0; i < workbook.Sheets.Count; i++)
        {
            var sheet = workbook.Sheets[i];
            sb.Append($"    <style:style style:name=\"ta{i + 1}\" style:family=\"table\" style:master-page-name=\"Default\">\n");
            sb.Append("      <style:table-properties table:display=\"true\" style:writing-mode=\"lr-tb\"");
            if (!string.IsNullOrEmpty(sheet.TabColor))
                sb.Append($" table:tab-color=\"{ArgbToOds(sheet.TabColor)}\"");
            sb.Append("/>\n    </style:style>\n");
        }
        // Graphic styles for images and chart frames
        sb.Append("    <style:style style:name=\"gr1\" style:family=\"graphic\" style:parent-style-name=\"Default\">\n");
        sb.Append("      <style:graphic-properties draw:stroke=\"none\" draw:fill=\"none\" draw:ole-draw-aspect=\"1\"/>\n");
        sb.Append("    </style:style>\n");
        sb.Append("    <style:style style:name=\"gr2\" style:family=\"graphic\" style:parent-style-name=\"Default\">\n");
        sb.Append("      <style:graphic-properties draw:stroke=\"none\" draw:fill=\"none\"\n");
        sb.Append("        draw:color-mode=\"standard\" fo:clip=\"rect(0in,0in,0in,0in)\"\n");
        sb.Append("        draw:image-opacity=\"100%\" style:mirror=\"none\"/>\n");
        sb.Append("    </style:style>\n");
        sb.Append("  </office:automatic-styles>\n");

        // body → spreadsheet → tables
        sb.Append("  <office:body>\n    <office:spreadsheet>\n");
        for (int i = 0; i < workbook.Sheets.Count; i++)
            AppendSheet(sb, workbook.Sheets[i], i + 1, styles, numFmtLookup, colWidthStyles);
        sb.Append("    </office:spreadsheet>\n  </office:body>\n");
        sb.Append("</office:document-content>\n");

        return Encoding.UTF8.GetBytes(sb.ToString());
    }

    // ── Document open element ──────────────────────────────────────────────────

    private static void AppendDocumentOpen(StringBuilder sb)
    {
        sb.Append("<office:document-content\n");
        sb.Append("  xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"\n");
        sb.Append("  xmlns:ooo=\"http://openoffice.org/2004/office\"\n");
        sb.Append("  xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\"\n");
        sb.Append("  xmlns:xlink=\"http://www.w3.org/1999/xlink\"\n");
        sb.Append("  xmlns:dc=\"http://purl.org/dc/elements/1.1/\"\n");
        sb.Append("  xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\"\n");
        sb.Append("  xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"\n");
        sb.Append("  xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\"\n");
        sb.Append("  xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\"\n");
        sb.Append("  xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\"\n");
        sb.Append("  xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"\n");
        sb.Append("  xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\"\n");
        sb.Append("  xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\"\n");
        sb.Append("  xmlns:loext=\"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0\"\n");
        sb.Append("  office:version=\"1.4\">\n");
        sb.Append("  <office:scripts/>\n");
    }

    // ── XlsxFont face declarations ─────────────────────────────────────────────────

    private static void AppendFontFaces(StringBuilder sb, XlsxStyles styles)
    {
        sb.Append("  <office:font-face-decls>\n");
        sb.Append("    <style:font-face style:name=\"Liberation Sans\" svg:font-family=\"'Liberation Sans'\" style:font-family-generic=\"swiss\" style:font-pitch=\"variable\"/>\n");
        // Add any non-Calibri fonts used
        var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Liberation Sans" };
        foreach (var f in styles.Fonts)
        {
            if (!fonts.Contains(f.Name))
            {
                fonts.Add(f.Name);
                sb.Append($"    <style:font-face style:name=\"{XmlEsc(f.Name)}\" svg:font-family=\"{XmlEsc(f.Name)}\" style:font-family-generic=\"swiss\"/>\n");
            }
        }
        sb.Append("  </office:font-face-decls>\n");
    }

    // ── Number format styles ───────────────────────────────────────────────────

    private static void AppendNumberFormats(StringBuilder sb,
        SortedSet<int> usedXfIndices, XlsxStyles styles,
        Dictionary<int, string> numFmtLookup)
    {
        // Emit one number-format style per unique numFmtId used
        var emitted = new HashSet<int>();
        foreach (int xfIdx in usedXfIndices)
        {
            if (xfIdx < 0 || xfIdx >= styles.CellXfs.Count) continue;
            var xf = styles.CellXfs[xfIdx];
            if (xf.NumFmtId == 0) continue;
            if (!emitted.Add(xf.NumFmtId)) continue;

            string fmtCode = numFmtLookup.TryGetValue(xf.NumFmtId, out var c) ? c : "General";
            string styleName = $"NF{xf.NumFmtId}";
            AppendNumberFormatStyle(sb, styleName, fmtCode);
        }
    }

    private static void AppendNumberFormatStyle(StringBuilder sb, string name, string xlsxFmt)
    {
        switch (xlsxFmt)
        {
            // Integer / whole number
            case "0":
                sb.Append($"    <number:number-style style:name=\"{name}\">");
                sb.Append("<number:number number:decimal-places=\"0\" number:min-integer-digits=\"1\"/>");
                sb.Append("</number:number-style>\n");
                break;

            // Thousands no decimal
            case "#,##0":
                sb.Append($"    <number:number-style style:name=\"{name}\">");
                sb.Append("<number:number number:decimal-places=\"0\" number:min-integer-digits=\"1\" number:grouping=\"true\"/>");
                sb.Append("</number:number-style>\n");
                break;

            // Thousands 2 decimal
            case "#,##0.00":
                sb.Append($"    <number:number-style style:name=\"{name}\">");
                sb.Append("<number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/>");
                sb.Append("</number:number-style>\n");
                break;

            // Percentage 2dp
            case "0.00%":
                sb.Append($"    <number:percentage-style style:name=\"{name}\">");
                sb.Append("<number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/>");
                sb.Append("<number:text>%</number:text>");
                sb.Append("</number:percentage-style>\n");
                break;

            // Percentage 0dp
            case "0%":
                sb.Append($"    <number:percentage-style style:name=\"{name}\">");
                sb.Append("<number:number number:decimal-places=\"0\" number:min-decimal-places=\"0\" number:min-integer-digits=\"1\"/>");
                sb.Append("<number:text>%</number:text>");
                sb.Append("</number:percentage-style>\n");
                break;

            // Currency $X,XXX.XX
            case "\"$\"#,##0.00":
                sb.Append($"    <number:currency-style style:name=\"{name}\">");
                sb.Append("<number:currency-symbol>$</number:currency-symbol>");
                sb.Append("<number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/>");
                sb.Append("</number:currency-style>\n");
                break;

            // Accounting (simplified to currency)
            case var s when s.StartsWith("_($"):
                sb.Append($"    <number:currency-style style:name=\"{name}\">");
                sb.Append("<number:currency-symbol>$</number:currency-symbol>");
                sb.Append("<number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/>");
                sb.Append("</number:currency-style>\n");
                break;

            // Short date m/d/yyyy
            case "m/d/yyyy":
                sb.Append($"    <number:date-style style:name=\"{name}\">");
                sb.Append("<number:month/><number:text>/</number:text><number:day/><number:text>/</number:text><number:year number:style=\"long\"/>");
                sb.Append("</number:date-style>\n");
                break;

            // Long date
            case var s when s.Contains("F800") || s.Contains("dddd"):
                sb.Append($"    <number:date-style style:name=\"{name}\">");
                sb.Append("<number:day-of-week number:style=\"long\"/><number:text>, </number:text>");
                sb.Append("<number:month number:textual=\"true\" number:style=\"long\"/><number:text> </number:text>");
                sb.Append("<number:day number:style=\"long\"/><number:text>, </number:text>");
                sb.Append("<number:year number:style=\"long\"/>");
                sb.Append("</number:date-style>\n");
                break;

            // DateTime m/d/yy h:mm
            case "m/d/yy h:mm":
                sb.Append($"    <number:date-style style:name=\"{name}\">");
                sb.Append("<number:month/><number:text>/</number:text><number:day/><number:text>/</number:text><number:year/>");
                sb.Append("<number:text> </number:text><number:hours/><number:text>:</number:text><number:minutes number:style=\"long\"/>");
                sb.Append("</number:date-style>\n");
                break;

            // Time 12h h:mm:ss AM/PM
            case "h:mm:ss AM/PM":
                sb.Append($"    <number:time-style style:name=\"{name}\">");
                sb.Append("<number:hours/><number:text>:</number:text><number:minutes number:style=\"long\"/><number:text>:</number:text>");
                sb.Append("<number:seconds number:style=\"long\"/><number:text> </number:text><number:am-pm/>");
                sb.Append("</number:time-style>\n");
                break;

            // Time 24h h:mm:ss
            case "h:mm:ss":
                sb.Append($"    <number:time-style style:name=\"{name}\">");
                sb.Append("<number:hours/><number:text>:</number:text><number:minutes number:style=\"long\"/><number:text>:</number:text>");
                sb.Append("<number:seconds number:style=\"long\"/>");
                sb.Append("</number:time-style>\n");
                break;

            // Scientific
            case "0.00E+00":
                sb.Append($"    <number:number-style style:name=\"{name}\">");
                sb.Append("<number:scientific-number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:min-exponent-digits=\"2\" number:forced-exponent-sign=\"true\"/>");
                sb.Append("</number:number-style>\n");
                break;

            // Fraction # ??/??
            case "# ??/??":
                sb.Append($"    <number:number-style style:name=\"{name}\">");
                sb.Append("<number:fraction number:min-integer-digits=\"0\" number:min-numerator-digits=\"2\" number:min-denominator-digits=\"2\"/>");
                sb.Append("</number:number-style>\n");
                break;

            // US Phone (###) ###-####
            case "(###) ###-####":
                sb.Append($"    <number:number-style style:name=\"{name}\">");
                sb.Append("<number:number number:decimal-places=\"0\" number:min-integer-digits=\"0\">");
                sb.Append("<number:embedded-text number:position=\"7\">) </number:embedded-text>");
                sb.Append("<number:embedded-text number:position=\"4\">-</number:embedded-text>");
                sb.Append("</number:number>");
                sb.Append("</number:number-style>\n");
                break;

            // ZIP code 00000
            case "00000":
                sb.Append($"    <number:number-style style:name=\"{name}\">");
                sb.Append("<number:number number:decimal-places=\"0\" number:min-integer-digits=\"5\"/>");
                sb.Append("</number:number-style>\n");
                break;

            // Text/@
            case "@":
                // No number format needed — cell will use string type
                break;

            default:
                // Fallback: generic 2-decimal number
                sb.Append($"    <number:number-style style:name=\"{name}\">");
                sb.Append("<number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/>");
                sb.Append("</number:number-style>\n");
                break;
        }
    }

    // ── XlsxCell styles ────────────────────────────────────────────────────────────

    private static void AppendCellStyles(StringBuilder sb, SortedSet<int> usedXfIndices,
        XlsxStyles styles, Dictionary<int, string> numFmtLookup)
    {
        foreach (int xfIdx in usedXfIndices)
        {
            if (xfIdx < 0 || xfIdx >= styles.CellXfs.Count) continue;
            var xf   = styles.CellXfs[xfIdx];
            var font = xfIdx < styles.CellXfs.Count && xf.FontId < styles.Fonts.Count
                       ? styles.Fonts[xf.FontId] : XlsxFont.Default;
            var fill = xf.FillId < styles.Fills.Count ? styles.Fills[xf.FillId] : null;
            var border = xf.BorderId < styles.Borders.Count ? styles.Borders[xf.BorderId] : null;
            var align  = xf.XlsxAlignment;

            string? numFmtStyleRef = xf.NumFmtId != 0 ? $"NF{xf.NumFmtId}" : null;

            sb.Append($"    <style:style style:name=\"ce{xfIdx}\" style:family=\"table-cell\" style:parent-style-name=\"Default\"");
            if (numFmtStyleRef != null)
                sb.Append($" style:data-style-name=\"{numFmtStyleRef}\"");
            sb.Append(">\n");

            // XlsxCell properties (background, border, wrap)
            bool hasCellProps = (fill?.PatternType == "solid" && fill.FgColor != null) ||
                                 (border != null && HasBorder(border)) ||
                                 align?.WrapText == true;
            if (hasCellProps)
            {
                sb.Append("      <style:table-cell-properties");
                if (fill?.PatternType == "solid" && fill.FgColor != null)
                    sb.Append($" fo:background-color=\"{ArgbToOds(fill.FgColor)}\"");
                if (border != null && HasBorder(border))
                {
                    if (AllBorderSidesEqual(border))
                    {
                        sb.Append($" fo:border=\"{OdsBorderValue(border.Left)}\"");
                    }
                    else
                    {
                        if (border.Left.Style   != null) sb.Append($" fo:border-left=\"{OdsBorderValue(border.Left)}\"");
                        if (border.Right.Style  != null) sb.Append($" fo:border-right=\"{OdsBorderValue(border.Right)}\"");
                        if (border.Top.Style    != null) sb.Append($" fo:border-top=\"{OdsBorderValue(border.Top)}\"");
                        if (border.Bottom.Style != null) sb.Append($" fo:border-bottom=\"{OdsBorderValue(border.Bottom)}\"");
                    }
                }
                if (align?.WrapText == true)
                    sb.Append(" fo:wrap-option=\"wrap\"");
                sb.Append(" style:rotation-align=\"none\"/>\n");
            }

            // Text properties (font, size, bold, italic, color, underline, strikethrough)
            bool hasTextProps = font.Bold || font.Italic || font.Underline || font.Strikethrough ||
                                font.Name != "Liberation Sans" || Math.Abs(font.Size - 10) > 0.01 ||
                                font.RgbColor != null;
            if (hasTextProps)
            {
                sb.Append("      <style:text-properties");
                if (font.Name != "Liberation Sans")
                    sb.Append($" style:font-name=\"{XmlEsc(font.Name)}\"");
                if (Math.Abs(font.Size - 10) > 0.01)
                    sb.Append($" fo:font-size=\"{font.Size.ToString("G", CultureInfo.InvariantCulture)}pt\"");
                if (font.RgbColor != null)
                    sb.Append($" fo:color=\"{ArgbToOds(font.RgbColor)}\"");
                if (font.Bold)
                    sb.Append(" fo:font-weight=\"bold\"");
                if (font.Italic)
                    sb.Append(" fo:font-style=\"italic\"");
                if (font.Underline)
                    sb.Append(" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\"");
                if (font.Strikethrough)
                    sb.Append(" style:text-line-through-style=\"solid\"");
                sb.Append("/>\n");
            }

            // Paragraph properties (alignment)
            if (align != null && (align.Horizontal != null || align.Vertical != null))
            {
                sb.Append("      <style:paragraph-properties");
                if (align.Horizontal != null)
                {
                    string odsAlign = align.Horizontal switch
                    {
                        "right"   => "end",
                        "center"  => "center",
                        "justify" => "justify",
                        _         => "start",
                    };
                    sb.Append($" fo:text-align=\"{odsAlign}\"");
                }
                sb.Append("/>\n");
            }

            sb.Append("    </style:style>\n");
        }
    }

    // ── Column styles ──────────────────────────────────────────────────────────

    private static void AppendColumnStyles(StringBuilder sb,
        Dictionary<double, string> colWidthStyles)
    {
        foreach (var (width, name) in colWidthStyles)
        {
            double inches = width < 0 ? 0.889 : width * ColWidthPerChar;
            sb.Append($"    <style:style style:name=\"{name}\" style:family=\"table-column\">\n");
            sb.Append($"      <style:table-column-properties fo:break-before=\"auto\" style:column-width=\"{inches:F4}in\"/>\n");
            sb.Append("    </style:style>\n");
        }
    }

    // ── Sheet (table) output ───────────────────────────────────────────────────

    private static void AppendSheet(StringBuilder sb, XlsxWorksheet sheet, int sheetNum,
        XlsxStyles styles, Dictionary<int, string> numFmtLookup,
        Dictionary<double, string> colWidthStyles)
    {
        // Derive sheet-level default cell style name (null when no SheetStyle is set)
        string? defaultCellStyle = sheet.SheetStyle != null ? $"sh{sheetNum}" : null;

        sb.Append($"      <table:table table:name=\"{XmlEsc(sheet.Name)}\" table:style-name=\"ta{sheetNum}\">\n");

        // Column definitions
        AppendColumns(sb, sheet, colWidthStyles, defaultCellStyle);

        // Determine max row with data
        int maxRow = sheet.Rows.Count > 0
            ? sheet.Rows.Max(r => r.RowIndex)
            : 0;

        // Track which rows have images (to inject draw:frame in that row)
        // Key = fromRow (0-based), Value = list of images
        var imagesByRow = new Dictionary<int, List<XlsxImageEntry>>();
        foreach (var img in sheet.Images)
        {
            int key = img.Anchor?.FromRow ?? 0;
            if (!imagesByRow.TryGetValue(key, out var lst))
                imagesByRow[key] = lst = [];
            lst.Add(img);
        }

        // Track which rows have charts
        var chartsByRow = new Dictionary<int, List<ChartInfo>>();
        foreach (var c in sheet.Charts)
        {
            int key = c.Anchor?.FromRow ?? 0;
            if (!chartsByRow.TryGetValue(key, out var lst))
                chartsByRow[key] = lst = [];
            lst.Add(c);
        }

        // Compute overall max row (including image/chart anchors)
        foreach (int r in imagesByRow.Keys)
            maxRow = Math.Max(maxRow, r + 1);
        foreach (int r in chartsByRow.Keys)
            maxRow = Math.Max(maxRow, r + 1);

        // Extend written rows so the sheet-default background style is visible
        // below the data area (table:default-cell-style-name only applies to
        // rows that are actually written to the XML).
        if (sheet.SheetStyle != null)
            maxRow = Math.Max(maxRow, 50);

        // Build a row-indexed lookup for fast access
        var rowByIndex = new Dictionary<int, XlsxRow>();
        foreach (var row in sheet.Rows)
            rowByIndex[row.RowIndex] = row;

        // Write rows 1 → maxRow
        int prevRow = 0;
        for (int ri = 1; ri <= maxRow; ri++)
        {
            // Write blank rows before this one if needed
            if (ri - prevRow > 1)
            {
                int gap = ri - prevRow - 1;
                // Blank cells inherit their background from the column's default-cell-style-name
                if (gap == 1)
                    sb.Append("      <table:table-row table:style-name=\"ro1\"><table:table-cell table:number-columns-repeated=\"16384\"/></table:table-row>\n");
                else
                    sb.Append($"      <table:table-row table:style-name=\"ro1\" table:number-rows-repeated=\"{gap}\"><table:table-cell table:number-columns-repeated=\"16384\"/></table:table-row>\n");
            }
            prevRow = ri;

            sb.Append("      <table:table-row table:style-name=\"ro1\">\n");

            // Images/charts anchored at this row (0-based → row ri means 0-based = ri-1)
            bool hasFloating = imagesByRow.ContainsKey(ri - 1) || chartsByRow.ContainsKey(ri - 1);

            if (rowByIndex.TryGetValue(ri, out var rowModel) && rowModel.Cells.Count > 0)
            {
                AppendRowCells(sb, rowModel, styles, numFmtLookup, sheet.Name,
                               imagesByRow, chartsByRow, ri - 1, defaultCellStyle, sheet.ColWidths);
            }
            else if (hasFloating)
            {
                // Row has no data cells but has images/charts — emit one cell with the frame
                sb.Append("        <table:table-cell>\n");
                EmitFloating(sb, imagesByRow, chartsByRow, ri - 1, sheet.Name, sheet.ColWidths);
                sb.Append("        </table:table-cell>\n");
                sb.Append("        <table:table-cell table:number-columns-repeated=\"16383\"/>\n");
            }
            else
            {
                sb.Append("        <table:table-cell table:number-columns-repeated=\"16384\"/>\n");
            }

            sb.Append("      </table:table-row>\n");
        }

        sb.Append("      </table:table>\n");
    }

    // ── Column declarations ────────────────────────────────────────────────────

    private static void AppendColumns(StringBuilder sb, XlsxWorksheet sheet,
        Dictionary<double, string> colWidthStyles, string? defaultCellStyle)
    {
        string dcsa = defaultCellStyle != null
            ? $" table:default-cell-style-name=\"{defaultCellStyle}\""
            : "";

        if (sheet.ColWidths.Count == 0)
        {
            // All columns use default width
            sb.Append($"        <table:table-column table:style-name=\"co0\"{dcsa} table:number-columns-repeated=\"16384\"/>\n");
            return;
        }

        // Build a sparse list of column widths
        // ColWidths stores (min, max, width) ranges (1-based)
        var widthByCol = new Dictionary<int, double>();
        foreach (var (min, max, w) in sheet.ColWidths)
            for (int c = min; c <= max; c++)
                widthByCol[c] = w;

        int maxCol = widthByCol.Keys.Count > 0 ? widthByCol.Keys.Max() : 0;
        for (int c = 1; c <= maxCol; c++)
        {
            if (!widthByCol.ContainsKey(c))
                widthByCol[c] = -1; // default
        }

        // Group consecutive same-width columns
        double curWidth = widthByCol.TryGetValue(1, out var w0) ? w0 : -1;
        int    curStart = 1;
        for (int c = 2; c <= maxCol + 1; c++)
        {
            double w = c <= maxCol && widthByCol.TryGetValue(c, out var ww) ? ww : -2;
            if (w != curWidth || c > maxCol)
            {
                string sn = colWidthStyles.TryGetValue(curWidth, out var sn2) ? sn2 : "co0";
                int count = c - curStart;
                if (count == 1)
                    sb.Append($"        <table:table-column table:style-name=\"{sn}\"{dcsa}/>\n");
                else
                    sb.Append($"        <table:table-column table:style-name=\"{sn}\"{dcsa} table:number-columns-repeated=\"{count}\"/>\n");
                curWidth = w;
                curStart = c;
            }
        }
        // Remaining columns (default)
        if (maxCol < 16384)
            sb.Append($"        <table:table-column table:style-name=\"co0\"{dcsa} table:number-columns-repeated=\"{16384 - maxCol}\"/>\n");
    }

    // ── XlsxRow cells ──────────────────────────────────────────────────────────────

    private static void AppendRowCells(StringBuilder sb, XlsxRow rowModel,
        XlsxStyles styles, Dictionary<int, string> numFmtLookup, string sheetName,
        Dictionary<int, List<XlsxImageEntry>> imagesByRow,
        Dictionary<int, List<ChartInfo>> chartsByRow,
        int rowIndex0, string? defaultCellStyle,
        IReadOnlyList<(int Min, int Max, double Width)> colWidths)
    {
        // Sort cells by column
        var cells = rowModel.Cells.OrderBy(c => CellReference.Parse(c.Reference).col).ToList();
        int prevCol = 0;

        for (int ci = 0; ci < cells.Count; ci++)
        {
            var cell  = cells[ci];
            var (_, col) = CellReference.Parse(cell.Reference);

            // Empty cells between previous and this one (inherit from column default-cell-style-name)
            int gap = col - prevCol - 1;
            if (gap == 1)
                sb.Append("        <table:table-cell/>\n");
            else if (gap > 1)
                sb.Append($"        <table:table-cell table:number-columns-repeated=\"{gap}\"/>\n");
            prevCol = col;

            // Is this the first cell? If so, attach floating objects anchored to this row
            bool injectFloating = ci == 0;

            AppendCell(sb, cell, styles, numFmtLookup, sheetName, defaultCellStyle);

            // Inject floating objects (images, charts) anchored to this row into the first cell
            if (injectFloating && (imagesByRow.ContainsKey(rowIndex0) || chartsByRow.ContainsKey(rowIndex0)))
            {
                // We need to close the cell and re-open with the frame inside
                // Actually, draw:frame goes INSIDE the table:table-cell element.
                // Since AppendCell already closed the element, we need a different approach.
                // → The floating injection must happen before the cell is closed.
                // Workaround: inject into the NEXT empty cell position by inserting a
                //             standalone cell with the frame(s).
                // We'll handle this by emitting a separate cell after the data cells.
                // (see EmitFloating call below after the loop)
            }
        }

        // Emit floating objects in a cell after the last data cell (if any)
        if (imagesByRow.ContainsKey(rowIndex0) || chartsByRow.ContainsKey(rowIndex0))
        {
            sb.Append("        <table:table-cell>\n");
            EmitFloating(sb, imagesByRow, chartsByRow, rowIndex0, sheetName, colWidths);
            sb.Append("        </table:table-cell>\n");
        }

        // Trailing empty columns
        if (prevCol < 16384)
        {
            int trail = 16384 - prevCol;
            sb.Append($"        <table:table-cell table:number-columns-repeated=\"{trail}\"/>\n");
        }
    }

    // ── Single cell ────────────────────────────────────────────────────────────

    private static void AppendCell(StringBuilder sb, XlsxCell cell,
        XlsxStyles styles, Dictionary<int, string> numFmtLookup, string sheetName,
        string? defaultCellStyle)
    {
        // Cells with no explicit style use the sheet default (if set) so the background
        // inherits correctly. Omitting table:style-name lets the column's
        // table:default-cell-style-name take effect — but we set it explicitly here so
        // that data cells (which have a table:table-cell element) also apply the sheet style.
        string styleName = cell.StyleIndex.HasValue
            ? $"ce{cell.StyleIndex.Value}"
            : (defaultCellStyle ?? "Default");
        string? fmtCode  = null;
        if (cell.StyleIndex.HasValue && cell.StyleIndex.Value < styles.CellXfs.Count)
        {
            int numFmtId = styles.CellXfs[cell.StyleIndex.Value].NumFmtId;
            numFmtLookup.TryGetValue(numFmtId, out fmtCode);
        }

        if (cell.Value is null)
        {
            sb.Append($"        <table:table-cell table:style-name=\"{styleName}\"/>\n");
            return;
        }

        switch (cell.Value)
        {
            case CellValue.Text t:
                sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" office:value-type=\"string\" calcext:value-type=\"string\">\n");
                sb.Append($"          <text:p>{XmlEsc(t.Value)}</text:p>\n");
                sb.Append("        </table:table-cell>\n");
                break;

            case CellValue.Boolean b:
                string bVal  = b.Value ? "true" : "false";
                string bText = b.Value ? "TRUE" : "FALSE";
                sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" office:value-type=\"boolean\" office:boolean-value=\"{bVal}\" calcext:value-type=\"boolean\">\n");
                sb.Append($"          <text:p>{bText}</text:p>\n");
                sb.Append("        </table:table-cell>\n");
                break;

            case CellValue.Number n:
                AppendNumberCell(sb, n.Value, fmtCode, styleName);
                break;

            case CellValue.Date d:
                AppendDateCell(sb, d.Value, fmtCode, styleName);
                break;

            case CellValue.Formula f:
                AppendFormulaCell(sb, f, styleName, sheetName);
                break;

            case CellValue.Error e:
                sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" table:formula=\"{e.Code.ToOdsFormula()}\" office:value-type=\"string\" calcext:value-type=\"error\">\n");
                sb.Append($"          <text:p>{XmlEsc(e.Code.ToXmlString())}</text:p>\n");
                sb.Append("        </table:table-cell>\n");
                break;

            default:
                sb.Append($"        <table:table-cell table:style-name=\"{styleName}\"/>\n");
                break;
        }
    }

    private static void AppendNumberCell(StringBuilder sb, double value,
        string? fmtCode, string styleName)
    {
        bool isPct = fmtCode != null &&
                     (fmtCode == "0.00%" || fmtCode == "0%" ||
                      (fmtCode.EndsWith('%') && !fmtCode.Contains("[-")));

        if (isPct)
        {
            string display = FormatPct(value, fmtCode!);
            sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" office:value-type=\"percentage\" office:value=\"{V(value)}\" calcext:value-type=\"percentage\">\n");
            sb.Append($"          <text:p>{XmlEsc(display)}</text:p>\n");
            sb.Append("        </table:table-cell>\n");
        }
        else
        {
            string display = FormatNumber(value, fmtCode);
            sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" office:value-type=\"float\" office:value=\"{V(value)}\" calcext:value-type=\"float\">\n");
            sb.Append($"          <text:p>{XmlEsc(display)}</text:p>\n");
            sb.Append("        </table:table-cell>\n");
        }
    }

    private static void AppendDateCell(StringBuilder sb, DateTime dt,
        string? fmtCode, string styleName)
    {
        bool isTimeOnly = fmtCode != null &&
                          (fmtCode == "h:mm:ss AM/PM" || fmtCode == "h:mm:ss") &&
                          !fmtCode.Contains("yyyy") && !fmtCode.Contains("yy");

        bool isDateTime = fmtCode != null &&
                          (fmtCode.Contains("h:mm") || fmtCode.Contains("h:ss")) &&
                          (fmtCode.Contains("yy") || fmtCode.Contains("d"));

        if (isTimeOnly)
        {
            // Time only → time value-type
            var ts = dt.TimeOfDay;
            string iso = $"PT{(int)ts.TotalHours:D2}H{ts.Minutes:D2}M{ts.Seconds:D2}S";
            string display = fmtCode!.Contains("AM/PM")
                ? dt.ToString("h:mm:ss tt", CultureInfo.InvariantCulture)
                : dt.ToString("H:mm:ss", CultureInfo.InvariantCulture);
            sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" office:value-type=\"time\" office:time-value=\"{iso}\" calcext:value-type=\"time\">\n");
            sb.Append($"          <text:p>{XmlEsc(display)}</text:p>\n");
            sb.Append("        </table:table-cell>\n");
        }
        else if (isDateTime)
        {
            string iso     = dt.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
            string display = dt.ToString("M/d/yy H:mm", CultureInfo.InvariantCulture);
            sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" office:value-type=\"date\" office:date-value=\"{iso}\" calcext:value-type=\"date\">\n");
            sb.Append($"          <text:p>{XmlEsc(display)}</text:p>\n");
            sb.Append("        </table:table-cell>\n");
        }
        else
        {
            // Date only
            string iso     = dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            string display = fmtCode != null && (fmtCode.Contains("F800") || fmtCode.Contains("dddd"))
                ? dt.ToString("dddd, MMMM dd, yyyy", CultureInfo.InvariantCulture)
                : dt.ToString("M/d/yyyy", CultureInfo.InvariantCulture);
            sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" office:value-type=\"date\" office:date-value=\"{iso}\" calcext:value-type=\"date\">\n");
            sb.Append($"          <text:p>{XmlEsc(display)}</text:p>\n");
            sb.Append("        </table:table-cell>\n");
        }
    }

    private static void AppendFormulaCell(StringBuilder sb, CellValue.Formula f,
        string styleName, string sheetName)
    {
        string odsFormula = ConvertFormula(f.Expression, sheetName);
        sb.Append($"        <table:table-cell table:style-name=\"{styleName}\" table:formula=\"{XmlEsc(odsFormula)}\"");
        // Add default value-type so readers that don't recalc show something
        sb.Append(" office:value-type=\"float\" office:value=\"0\"");
        sb.Append(">\n");
        sb.Append("          <text:p/>\n");
        sb.Append("        </table:table-cell>\n");
    }

    // ── Floating objects (images and charts) ──────────────────────────────────

    private static void EmitFloating(StringBuilder sb,
        Dictionary<int, List<XlsxImageEntry>> imagesByRow,
        Dictionary<int, List<ChartInfo>> chartsByRow,
        int rowIndex0, string sheetName,
        IReadOnlyList<(int Min, int Max, double Width)> colWidths)
    {
        // Images
        if (imagesByRow.TryGetValue(rowIndex0, out var images))
        {
            foreach (var img in images)
            {
                var anchor = img.Anchor ?? new ObjectAnchor();
                double x      = ComputeColSpanWidth(0, anchor.FromCol, colWidths);
                double y      = 0;
                double width  = ComputeColSpanWidth(anchor.FromCol, anchor.ToCol, colWidths);
                double height = Math.Max(1, anchor.ToRow - anchor.FromRow) * DefaultRowH;
                string endCol  = ColLetter(anchor.ToCol + 1);
                string endCell = $"{XmlEsc(sheetName)}.{endCol}{anchor.ToRow + 1}";
                string imgPath = $"Pictures/image{img.MediaId}.{img.Extension}";

                sb.Append($"          <draw:frame draw:style-name=\"gr2\" draw:name=\"Image{img.MediaId}\"\n");
                sb.Append($"            table:end-cell-address=\"{endCell}\"\n");
                sb.Append($"            table:end-x=\"0in\" table:end-y=\"0in\"\n");
                sb.Append($"            svg:width=\"{width:F4}in\" svg:height=\"{height:F4}in\"\n");
                sb.Append($"            svg:x=\"{x:F4}in\" svg:y=\"{y:F4}in\">\n");
                sb.Append($"            <draw:image xlink:href=\"{imgPath}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\" draw:mime-type=\"{img.ContentType}\"/>\n");
                sb.Append("          </draw:frame>\n");
            }
        }

        // Charts
        if (chartsByRow.TryGetValue(rowIndex0, out var charts))
        {
            foreach (var c in charts)
            {
                var anchor = c.Anchor ?? new ObjectAnchor();
                double x      = ComputeColSpanWidth(0, anchor.FromCol, colWidths);
                double y      = 0;
                double width  = ComputeColSpanWidth(anchor.FromCol, anchor.ToCol, colWidths);
                double height = Math.Max(1, anchor.ToRow - anchor.FromRow) * DefaultRowH;
                string endCol  = ColLetter(anchor.ToCol + 1);
                string endCell = $"{XmlEsc(sheetName)}.{endCol}{anchor.ToRow + 1}";

                sb.Append($"          <draw:frame draw:style-name=\"gr1\" draw:name=\"Object{c.ChartId}\"\n");
                sb.Append($"            table:end-cell-address=\"{endCell}\"\n");
                sb.Append($"            table:end-x=\"0in\" table:end-y=\"0in\"\n");
                sb.Append($"            svg:width=\"{width:F4}in\" svg:height=\"{height:F4}in\"\n");
                sb.Append($"            svg:x=\"{x:F4}in\" svg:y=\"{y:F4}in\">\n");
                sb.Append($"            <draw:object xlink:href=\"./Object {c.ChartId}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/>\n");
                sb.Append("          </draw:frame>\n");
            }
        }
    }

    /// <summary>
    /// Sums the actual column widths (in inches) for columns fromCol0..toCol0-1 (0-based indices).
    /// Falls back to the Excel default width (8.43 chars) for columns with no explicit width.
    /// </summary>
    private static double ComputeColSpanWidth(int fromCol0, int toCol0,
        IReadOnlyList<(int Min, int Max, double Width)> colWidths)
    {
        if (toCol0 <= fromCol0) return 0;

        // Build a 1-based column → width lookup from the ColWidths ranges
        var widthMap = new Dictionary<int, double>();
        foreach (var (min, max, w) in colWidths)
            for (int c = min; c <= max; c++)
                widthMap[c] = w;

        double inches = 0;
        for (int col0 = fromCol0; col0 < toCol0; col0++)
        {
            double chars = widthMap.TryGetValue(col0 + 1, out var cw) ? cw : 8.43;
            inches += chars * ColWidthPerChar;
        }
        return Math.Max(0.25, inches);
    }

    // ── Sheet-level default cell styles ──────────────────────────────────────

    /// <summary>
    /// Emits one ODS automatic cell style (<c>sh{N}</c>) per sheet that has a
    /// <see cref="SheetStyle"/> set.  These styles are referenced via
    /// <c>table:default-cell-style-name</c> on column definitions and on unstyled
    /// data cells so that empty cells also display the sheet background.
    /// </summary>
    private static void AppendSheetDefaultStyles(StringBuilder sb,
        IReadOnlyList<XlsxWorksheet> sheets)
    {
        for (int i = 0; i < sheets.Count; i++)
        {
            var ss = sheets[i].SheetStyle;
            if (ss is null) continue;

            sb.Append($"    <style:style style:name=\"sh{i + 1}\" style:family=\"table-cell\" style:parent-style-name=\"Default\">\n");

            // Background color
            if (ss.BackgroundColor != null)
            {
                sb.Append("      <style:table-cell-properties");
                sb.Append($" fo:background-color=\"{ArgbToOds(ss.BackgroundColor)}\"");
                sb.Append(" style:rotation-align=\"none\"/>\n");
            }

            // Font properties
            bool hasText = ss.FontName != null || ss.FontSize.HasValue ||
                           ss.FontColor != null || ss.Bold || ss.Italic;
            if (hasText)
            {
                sb.Append("      <style:text-properties");
                if (ss.FontName != null)
                    sb.Append($" style:font-name=\"{XmlEsc(ss.FontName)}\"");
                if (ss.FontSize.HasValue)
                    sb.Append($" fo:font-size=\"{ss.FontSize.Value.ToString("G", CultureInfo.InvariantCulture)}pt\"");
                if (ss.FontColor != null)
                    sb.Append($" fo:color=\"{ArgbToOds(ss.FontColor)}\"");
                if (ss.Bold)
                    sb.Append(" fo:font-weight=\"bold\"");
                if (ss.Italic)
                    sb.Append(" fo:font-style=\"italic\"");
                sb.Append("/>\n");
            }

            sb.Append("    </style:style>\n");
        }
    }

    // ── Value formatting helpers ───────────────────────────────────────────────

    private static string FormatNumber(double value, string? fmtCode)
    {
        if (fmtCode is null or "General" or "")
            return value.ToString("G", CultureInfo.InvariantCulture);

        return fmtCode switch
        {
            "0"           => ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture),
            "#,##0"       => value.ToString("#,##0", CultureInfo.InvariantCulture),
            "#,##0.00"    => value.ToString("#,##0.00", CultureInfo.InvariantCulture),
            "\"$\"#,##0.00" => "$" + Math.Abs(value).ToString("#,##0.00", CultureInfo.InvariantCulture),
            _ when fmtCode.StartsWith("_($")
                          => "$" + Math.Abs(value).ToString("#,##0.00", CultureInfo.InvariantCulture),
            "0.00E+00"    => value.ToString("0.00E+00", CultureInfo.InvariantCulture),
            "# ??/??"     => ToFraction(value),
            "(###) ###-####" => FormatPhone(value),
            "00000"       => ((long)Math.Round(value)).ToString("D5", CultureInfo.InvariantCulture),
            "@"           => value.ToString(CultureInfo.InvariantCulture),
            _             => value.ToString("G", CultureInfo.InvariantCulture),
        };
    }

    private static string FormatPct(double value, string fmtCode)
    {
        double pct = value * 100.0;
        return fmtCode == "0%"
            ? Math.Round(pct).ToString("0", CultureInfo.InvariantCulture) + "%"
            : pct.ToString("0.00", CultureInfo.InvariantCulture) + "%";
    }

    private static string ToFraction(double v)
    {
        int whole = (int)Math.Truncate(v);
        double frac = Math.Abs(v - whole);
        if (frac < 0.001) return whole.ToString();
        // Simple approximation: find nearest /2 /4 /8 /16 /32
        int bestN = 1, bestD = 2;
        double bestErr = double.MaxValue;
        for (int d = 2; d <= 32; d *= 2)
        {
            int n = (int)Math.Round(frac * d);
            double err = Math.Abs(frac - (double)n / d);
            if (err < bestErr) { bestErr = err; bestN = n; bestD = d; }
        }
        return whole == 0 ? $"{bestN}/{bestD}" : $"{whole} {bestN}/{bestD}";
    }

    private static string FormatPhone(double value)
    {
        long digits = (long)Math.Round(value);
        string s = digits.ToString("D10");
        if (s.Length >= 10)
            return $"({s[..3]}) {s[3..6]}-{s[6..10]}";
        return value.ToString(CultureInfo.InvariantCulture);
    }

    // ── Formula conversion (XLSX → ODS) ───────────────────────────────────────

    private static string ConvertFormula(string xlsxExpr, string defaultSheet)
    {
        // Convert cell references:  A1 → [.A1]  Sheet1!A1 → [Sheet1.A1]
        // Range:  A1:B10 → [.A1:.B10]  Sheet1!A1:B10 → [Sheet1.A1:Sheet1.B10]
        var sb = new StringBuilder("of:=");
        // Replace Sheet!$A$1:$B$10 patterns
        string result = System.Text.RegularExpressions.Regex.Replace(
            xlsxExpr,
            @"(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?",
            m =>
            {
                string sheet = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                string fromC = m.Groups[3].Value;
                string fromR = m.Groups[4].Value;
                if (m.Groups[5].Success)
                {
                    string toC = m.Groups[5].Value;
                    string toR = m.Groups[6].Value;
                    return $"[$'{sheet}'.${fromC}${fromR}:$'{sheet}'.${toC}${toR}]";
                }
                return $"[$'{sheet}'.${fromC}${fromR}]";
            });

        // Replace remaining bare cell ranges: A1:B10
        result = System.Text.RegularExpressions.Regex.Replace(
            result,
            @"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)",
            m => $"[.${m.Groups[1].Value}${m.Groups[2].Value}:.${m.Groups[3].Value}${m.Groups[4].Value}]");

        // Replace remaining bare cell refs: A1
        result = System.Text.RegularExpressions.Regex.Replace(
            result,
            @"(?<![.$\['])([A-Z]+)(\d+)(?![\]'])",
            m => $"[.${m.Groups[1].Value}${m.Groups[2].Value}]");

        return "of:=" + result;
    }

    // ── Utility helpers ────────────────────────────────────────────────────────

    /// <summary>Converts ARGB color "FFRRGGBB" to ODS "#rrggbb".</summary>
    private static string ArgbToOds(string argb)
    {
        if (argb.Length == 8)
            return "#" + argb[2..].ToLowerInvariant();
        if (argb.Length == 6)
            return "#" + argb.ToLowerInvariant();
        return "#000000";
    }

    /// <summary>Converts a 1-based column index to letter(s): 1→A, 27→AA.</summary>
    private static string ColLetter(int col)
    {
        var s = new StringBuilder();
        while (col > 0)
        {
            col--;
            s.Insert(0, (char)('A' + col % 26));
            col /= 26;
        }
        return s.ToString();
    }

    /// <summary>Formats a double for an XML attribute (no trailing zeros).</summary>
    private static string V(double d)
        => d.ToString("G15", CultureInfo.InvariantCulture);

    private static string XmlEsc(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
         .Replace("\"", "&quot;").Replace("'", "&apos;");

    private static bool HasBorder(XlsxBorder b) =>
        b.Left.Style != null || b.Right.Style != null ||
        b.Top.Style  != null || b.Bottom.Style != null;

    private static bool AllBorderSidesEqual(XlsxBorder b) =>
        b.Left.Style  == b.Right.Style  && b.Left.Color  == b.Right.Color  &&
        b.Left.Style  == b.Top.Style    && b.Left.Color  == b.Top.Color    &&
        b.Left.Style  == b.Bottom.Style && b.Left.Color  == b.Bottom.Color;

    private static string OdsBorderValue(XlsxBorderSide s)
    {
        string thickness = s.Style switch
        {
            "medium" or "mediumDashed" or "mediumDashDot" or "mediumDashDotDot" => "1.76pt",
            "thick"                                                               => "2.65pt",
            "hair"                                                                => "0.26pt",
            _                                                                     => "0.74pt",
        };
        string lineStyle = s.Style switch
        {
            "dashed" or "mediumDashed" or "dashDot" or "mediumDashDot"
                or "dashDotDot" or "mediumDashDotDot" or "slantDashDot" => "dashed",
            "dotted" or "hair"                                           => "dotted",
            "double"                                                     => "double",
            _                                                            => "solid",
        };
        string color = s.Color != null ? ArgbToOds(s.Color) : "#000000";
        return $"{thickness} {lineStyle} {color}";
    }

    /// <summary>Builds a reverse lookup from NumFmtId → format code string.</summary>
    private static Dictionary<int, string> BuildNumFmtLookup(XlsxStyles styles)
    {
        // Start with built-in IDs
        var lookup = new Dictionary<int, string>();
        foreach (var (code, id) in XlsxDataTypeFormats.BuiltInIds)
            lookup[id] = code;
        // Add custom ones
        foreach (var nf in styles.NumFmts)
            lookup[nf.NumFmtId] = nf.FormatCode;
        return lookup;
    }
}
