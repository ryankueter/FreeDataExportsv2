using System.Globalization;
using FreeDataExportsv2.Internal;

namespace FreeDataExportsv2
{
    /// <summary>One worksheet in an <see cref="XlsxFile"/>.</summary>
    public sealed class XlsxWorksheet
    {
        // ── Public ─────────────────────────────────────────────────────────────────

        public string  Name     { get; set; }
        /// <summary>ARGB hex tab color, e.g. "FFFF0000" for red. Null = default.</summary>
        public string? TabColor    { get; set; }
        /// <summary>Default background color, font, and size for the entire worksheet.</summary>
        public SheetStyle? SheetStyle { get; set; }

        // ── Internal (used by Parts) ───────────────────────────────────────────────

        internal int              SheetId          { get; }
        internal bool             IsTabSelected    { get; set; }
        internal string?          DimensionRef     { get; set; }
        internal double           DefaultRowHeight { get; set; } = 15;
        internal double           DyDescent        { get; set; } = 0.25;
        internal XlsxPageMargins XlsxPageMargins { get; set; } = new();
        internal List<XlsxRow>   Rows        { get; }     = [];
        internal List<(int Min, int Max, double Width)> ColWidths { get; } = [];
        internal List<XlsxTableInfo>  Tables    { get; } = [];
        internal List<ChartInfo>  Charts    { get; } = [];
        internal List<XlsxImageEntry> Images    { get; } = [];

        /// <summary>Global 1-based drawing ID assigned by XlsxFile before save. 0 = no drawing.</summary>
        internal int DrawingId       { get; set; }
        /// <summary>1-based relationship ID for the drawing within this sheet's rels file.</summary>
        internal int DrawingLocalRId { get; set; }

        internal XlsxStyles?                  _styles;
        internal XlsxErrorLog?                     _errorLog;
        internal Dictionary<DataType, string>? _formatOverrides;

        private int _nextRowIndex = 1;

        internal XlsxWorksheet(string name, int sheetId, XlsxStyles styles,
                                 XlsxErrorLog errorLog, Dictionary<DataType, string> formatOverrides,
                                 bool tabSelected = false)
        {
            Name             = name;
            SheetId          = sheetId;
            IsTabSelected    = tabSelected;
            _styles          = styles;
            _errorLog        = errorLog;
            _formatOverrides = formatOverrides;
        }

        // ── Public API ─────────────────────────────────────────────────────────────

        /// <summary>
        /// Appends a row at the next available index and returns a <see cref="XlsxRowBuilder"/>
        /// for fluent <c>AddCell</c> chaining.
        /// </summary>
        public XlsxRowBuilder AddRow()
        {
            int rowIndex = _nextRowIndex++;
            GetOrAddRow(rowIndex);
            return new XlsxRowBuilder(this, rowIndex);
        }

        /// <summary>
        /// Sets custom column widths (in character units).
        /// Returns <c>this</c> for chaining.
        /// </summary>
        public XlsxWorksheet ColumnWidths(params string[] widths)
        {
            ColWidths.Clear();
            for (int i = 0; i < widths.Length; i++)
            {
                if (double.TryParse(widths[i], NumberStyles.Any,
                        CultureInfo.InvariantCulture, out double w))
                    ColWidths.Add((i + 1, i + 1, w));
            }
            return this;
        }

        /// <summary>
        /// Sets the default background color, font, and size for the whole sheet.
        /// Returns <c>this</c> for chaining.
        /// </summary>
        public XlsxWorksheet ApplySheetStyle(SheetStyle style)
        {
            SheetStyle = style;
            return this;
        }

        // ── Internal: resolve SheetStyle to an XF style index ─────────────────────

        /// <summary>
        /// Converts <see cref="SheetStyle"/> to a style-table index that
        /// <see cref="XlsxWorksheetPart"/> can write on <c>&lt;col&gt;</c> and <c>&lt;row&gt;</c> elements.
        /// Returns <c>null</c> when no sheet style is configured.
        /// </summary>
        internal int? GetSheetStyleIndex()
        {
            if (SheetStyle is null || _styles is null) return null;

            var opts = new CellOptions
            {
                FontName        = SheetStyle.FontName,
                FontSize        = SheetStyle.FontSize,
                FontColor       = SheetStyle.FontColor,
                Bold            = SheetStyle.Bold,
                Italic          = SheetStyle.Italic,
                BackgroundColor = SheetStyle.BackgroundColor,
            };

            // Only call into XlsxStyles when at least one property is actually set
            bool hasFormatting =
                opts.FontName        is not null ||
                opts.FontSize.HasValue            ||
                opts.FontColor       is not null ||
                opts.Bold                         ||
                opts.Italic                       ||
                opts.BackgroundColor is not null;

            return hasFormatting
                ? _styles.GetOrAddCellXfFromOptions(opts, _formatOverrides)
                : null;
        }

        // ── Public SetCell API ────────────────────────────────────────────────────

        /// <summary>
        /// Sets a cell by A1-style reference (e.g. <c>"B3"</c>).
        /// Useful when you need to write to a specific cell without iterating rows.
        /// Returns <c>this</c> for chaining.
        /// </summary>
        public XlsxWorksheet SetCell(string cellRef, object? value,
                                       DataType dataType = DataType.General)
        {
            var (row, col) = CellReference.Parse(cellRef);
            SetCellCore(cellRef, row, col, value, dataType, null);
            return this;
        }

        /// <summary>
        /// Sets a cell by A1-style reference with full <see cref="CellOptions"/> (font, fill, etc.).
        /// Returns <c>this</c> for chaining.
        /// </summary>
        public XlsxWorksheet SetCell(string cellRef, object? value, CellOptions options)
        {
            var (row, col) = CellReference.Parse(cellRef);
            SetCellCore(cellRef, row, col, value, options.DataType, options);
            return this;
        }

        /// <summary>
        /// Sets a cell by 1-based row and column indices.
        /// Returns <c>this</c> for chaining.
        /// </summary>
        public XlsxWorksheet SetCell(int rowIndex, int colIndex, object? value,
                                       DataType dataType = DataType.General)
        {
            SetCellCore(CellReference.FromRowCol(rowIndex, colIndex),
                        rowIndex, colIndex, value, dataType, null);
            return this;
        }

        /// <summary>
        /// Sets a cell by 1-based row and column indices with full <see cref="CellOptions"/>.
        /// Returns <c>this</c> for chaining.
        /// </summary>
        public XlsxWorksheet SetCell(int rowIndex, int colIndex, object? value, CellOptions options)
        {
            SetCellCore(CellReference.FromRowCol(rowIndex, colIndex),
                        rowIndex, colIndex, value, options.DataType, options);
            return this;
        }

        // ── Table API ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Attaches an Excel table to this worksheet over the given cell range.
        /// If <see cref="XlsxTableDefinition.HasTotalsRow"/> is true and any columns have a
        /// <see cref="XlsxTotalsRowFunction"/>, SUBTOTAL formula cells are automatically added
        /// to the last row of the range.
        /// Returns <c>this</c> for chaining.
        /// </summary>
        /// <param name="cellRange">
        /// A1-style range that covers the entire table, including the header row and (if used)
        /// the totals row, e.g. <c>"A1:D26"</c>.
        /// </param>
        public XlsxWorksheet AddTable(string cellRange, XlsxTableDefinition definition)
        {
            var info = new XlsxTableInfo(cellRange, definition);
            Tables.Add(info);

            if (definition.HasTotalsRow)
                GenerateTotalsRow(cellRange, definition);

            return this;
        }

        // ── Image API ─────────────────────────────────────────────────────────────

        /// <summary>
        /// Embeds an image loaded from <paramref name="filePath"/> onto this worksheet.
        /// The content type is inferred from the file extension.
        /// </summary>
        /// <param name="filePath">Absolute or relative path to a PNG, JPEG, GIF, BMP, or TIFF file.</param>
        /// <param name="anchor">
        /// Optional two-cell anchor controlling position and size.
        /// Defaults to columns A–F, rows 2–14 (0-based).
        /// </param>
        public XlsxWorksheet AddImage(string filePath, ObjectAnchor? anchor = null)
        {
            var ext = Path.GetExtension(filePath).TrimStart('.').ToLowerInvariant();
            var contentType = ext switch
            {
                "jpg" or "jpeg" => "image/jpeg",
                "gif"           => "image/gif",
                "bmp"           => "image/bmp",
                "tiff" or "tif" => "image/tiff",
                "webp"          => "image/webp",
                _               => "image/png",
            };
            var bytes = File.ReadAllBytes(filePath);
            Images.Add(new XlsxImageEntry(bytes, contentType,
                        anchor ?? DefaultImageAnchor()));
            return this;
        }

        /// <summary>
        /// Embeds an image supplied as a raw byte array onto this worksheet.
        /// </summary>
        /// <param name="imageBytes">Raw image bytes (PNG, JPEG, GIF, BMP …).</param>
        /// <param name="contentType">MIME type, e.g. <c>"image/png"</c>. Default: <c>"image/png"</c>.</param>
        /// <param name="anchor">
        /// Optional two-cell anchor controlling position and size.
        /// Defaults to columns A–F, rows 2–14 (0-based).
        /// </param>
        public XlsxWorksheet AddImage(byte[] imageBytes, string contentType = "image/png",
                                        ObjectAnchor? anchor = null)
        {
            Images.Add(new XlsxImageEntry(imageBytes, contentType,
                        anchor ?? DefaultImageAnchor()));
            return this;
        }

        private static ObjectAnchor DefaultImageAnchor() =>
            new() { FromCol = 0, FromRow = 1, ToCol = 5, ToRow = 13 };

        // ── Chart API ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Embeds a chart on this worksheet.
        /// </summary>
        /// <param name="definition">The chart definition (type, series, title, legend).</param>
        /// <param name="anchor">
        /// Optional two-cell anchor controlling the chart's position and size.
        /// Defaults to columns A–H, rows 2–21 (0-based: FromCol=0, FromRow=1, ToCol=7, ToRow=20).
        /// </param>
        public XlsxWorksheet AddChart(ChartDefinition definition,
                                        ObjectAnchor? anchor = null)
        {
            Charts.Add(new ChartInfo(definition, anchor ?? new ObjectAnchor()));
            return this;
        }

        private void GenerateTotalsRow(string cellRange, XlsxTableDefinition definition)
        {
            var parts = cellRange.Split(':');
            if (parts.Length != 2) return;

            var (_, startCol) = CellReference.Parse(parts[0]);
            var (endRow, _)   = CellReference.Parse(parts[1]);

            for (int i = 0; i < definition.Columns.Count; i++)
            {
                var col  = definition.Columns[i];
                int code = SubtotalCode(col.TotalsFunction);
                if (code == 0) continue;

                var formula = $"SUBTOTAL({code},{definition.DisplayName}[{col.Name}])";
                var cellRef = CellReference.FromRowCol(endRow, startCol + i);
                var row     = GetOrAddRow(endRow);

                if (!row.Cells.Any(c => c.Reference == cellRef))
                    row.Cells.Add(new XlsxCell { Reference = cellRef,
                                                  Value     = CellValue.AsFormula(formula) });
            }
        }

        private static int SubtotalCode(XlsxTotalsRowFunction f) => f switch
        {
            XlsxTotalsRowFunction.Sum       => 109,
            XlsxTotalsRowFunction.Average   => 101,
            XlsxTotalsRowFunction.Count     => 103,
            XlsxTotalsRowFunction.CountNums => 102,
            XlsxTotalsRowFunction.Max       => 104,
            XlsxTotalsRowFunction.Min       => 105,
            XlsxTotalsRowFunction.StdDev    => 107,
            XlsxTotalsRowFunction.Var       => 110,
            _                           => 0,
        };

        // ── Core (SetCell + error-catching in one place) ───────────────────────────

        private XlsxCell SetCellCore(string cellRef, int rowIdx, int colIdx,
                                       object? value, DataType dataType, CellOptions? options)
        {
            try
            {
                var cv       = CoerceValue(value, dataType);
                var styleIdx = ResolveStyle(dataType, options, cv);
                return PlaceCell(cellRef, rowIdx, colIdx, cv, styleIdx);
            }
            catch (Exception ex)
            {
                _errorLog?.Add(Name, cellRef, value, ex);
                return SetErrorCell(cellRef, rowIdx, colIdx);
            }
        }

        private XlsxCell PlaceCell(string cellRef, int rowIdx, int colIdx,
                                     CellValue? cv, int? styleIdx)
        {
            var row  = GetOrAddRow(rowIdx);
            var cell = row.Cells.Find(c => c.Reference == cellRef);
            if (cell is null)
            {
                cell = new XlsxCell { Reference = cellRef };
                int ins = row.Cells.FindIndex(c => CellReference.Parse(c.Reference).col > colIdx);
                if (ins < 0) row.Cells.Add(cell);
                else         row.Cells.Insert(ins, cell);
            }
            cell.Value      = cv;
            cell.StyleIndex = styleIdx;
            return cell;
        }

        private XlsxCell SetErrorCell(string cellRef, int rowIdx, int colIdx)
        {
            try { return PlaceCell(cellRef, rowIdx, colIdx, null, _styles?.GetOrAddRedBorderXf()); }
            catch { return new XlsxCell { Reference = cellRef }; }
        }

        // ── Value coercion ─────────────────────────────────────────────────────────

        private static CellValue? CoerceValue(object? value, DataType dt)
        {
            if (value is null) return null;
            if (value is CellValue cv) return cv;

            if (dt is DataType.String or DataType.Text)
                return CellValue.Of(Convert.ToString(value) ?? string.Empty);

            if (dt == DataType.Boolean)
                return CellValue.Of(Convert.ToBoolean(value));

            if (dt == DataType.Formula)
            {
                var expr = (Convert.ToString(value) ?? string.Empty).TrimStart('=');
                return CellValue.AsFormula(expr);
            }

            if (dt == DataType.Error)
            {
                if (value is ErrorCode ec) return CellValue.AsError(ec);
                return CellValue.AsError(ErrorCode.InvalidValue);
            }

            if (XlsxDataTypeFormats.IsDateType(dt))
            {
                return value switch
                {
                    DateTime d  => CellValue.Of(d),
                    double   dv => CellValue.Of(System.DateTime.FromOADate(dv)),
                    _           => CellValue.Of(Convert.ToDateTime(value, CultureInfo.InvariantCulture)),
                };
            }

            // Natural numeric coercion
            return value switch
            {
                bool     b  => CellValue.Of(b ? 1.0 : 0.0),
                int      i  => CellValue.Of(i),
                long     l  => CellValue.Of(l),
                float    f  => CellValue.Of(f),
                double   d  => CellValue.Of(d),
                decimal  m  => CellValue.Of(m),
                DateTime dtt => CellValue.Of(dtt),
                string   s  => double.TryParse(s, NumberStyles.Any,
                                   CultureInfo.InvariantCulture, out var sd)
                               ? CellValue.Of(sd)
                               : dt == DataType.General
                                   ? CellValue.Of(s)
                                   : throw new InvalidCastException($"Cannot convert \"{s}\" to a numeric value for data type {dt}."),
                _ => CellValue.Of(Convert.ToString(value) ?? string.Empty),
            };
        }

        // ── Style resolution ───────────────────────────────────────────────────────

        private int? ResolveStyle(DataType dt, CellOptions? options, CellValue? cv)
        {
            if (_styles is null) return null;

            // Full CellOptions path (custom font / fill / alignment)
            if (options is not null && HasCustomFormatting(options))
                return _styles.GetOrAddCellXfFromOptions(options, _formatOverrides);

            // Auto-promote bare Date cells to ShortDate format
            if (dt == DataType.General && cv is CellValue.Date)
                dt = DataType.ShortDate;

            // Types that carry their own Excel cell-type attribute need no number format
            if (dt is DataType.String or DataType.Boolean or
                      DataType.Formula or DataType.Error or DataType.General)
                return null;

            var formatCode = XlsxDataTypeFormats.GetFormatCode(dt, _formatOverrides);
            return formatCode == "General" ? null : _styles.GetOrAddCellXf(formatCode);
        }

        private static bool HasCustomFormatting(CellOptions o) =>
            o.FontName is not null || o.FontSize.HasValue || o.FontColor is not null ||
            o.Bold || o.Italic || o.Underline || o.Strikethrough ||
            o.BackgroundColor is not null ||
            o.HorizontalAlign is not null || o.VerticalAlign is not null || o.WrapText ||
            o.BorderLeftStyle is not null || o.BorderRightStyle  is not null ||
            o.BorderTopStyle  is not null || o.BorderBottomStyle is not null;

        // ── XlsxRow management ─────────────────────────────────────────────────────────

        internal XlsxRow GetOrAddRow(int rowIndex)
        {
            _nextRowIndex = Math.Max(_nextRowIndex, rowIndex + 1);
            var row = Rows.Find(r => r.RowIndex == rowIndex);
            if (row is not null) return row;

            row = new XlsxRow { RowIndex = rowIndex };
            int ins = Rows.FindIndex(r => r.RowIndex > rowIndex);
            if (ins < 0) Rows.Add(row);
            else         Rows.Insert(ins, row);
            return row;
        }
    }

    // ── Fluent row builder ─────────────────────────────────────────────────────────

    /// <summary>
    /// Returned by <see cref="XlsxWorksheet.AddRow"/>; chain <c>AddCell</c> calls.
    /// </summary>
    public sealed class XlsxRowBuilder
    {
        private readonly XlsxWorksheet _sheet;
        private readonly int            _rowIndex;
        private int _nextCol = 1;

        internal XlsxRowBuilder(XlsxWorksheet sheet, int rowIndex)
        {
            _sheet    = sheet;
            _rowIndex = rowIndex;
        }

        /// <summary>Appends a cell with the given value and data type.</summary>
        public XlsxRowBuilder AddCell(object? value, DataType dataType = DataType.General)
        {
            _sheet.SetCell(_rowIndex, _nextCol++, value, dataType);
            return this;
        }

        /// <summary>Appends a cell with full <see cref="CellOptions"/> (font, fill, alignment, etc.).</summary>
        public XlsxRowBuilder AddCell(object? value, CellOptions options)
        {
            _sheet.SetCell(_rowIndex, _nextCol++, value, options);
            return this;
        }
    }
}
