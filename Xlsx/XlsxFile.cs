using System.IO.Compression;
using System.Text;
using FreeDataExportsv2.Internal;

namespace FreeDataExportsv2;

/// <summary>
/// Main entry point. Creates an .xlsx workbook.
/// </summary>
public sealed class XlsxFile
{
    // ── Core properties ────────────────────────────────────────────────────────

    public string   Creator        { get; set; } = string.Empty;
    public string   LastModifiedBy { get; set; } = string.Empty;
    public DateTime Created        { get; set; } = DateTime.UtcNow;
    public DateTime Modified       { get; set; } = DateTime.UtcNow;
    public string   Company        { get; set; } = string.Empty;
    public string   Application    { get; set; } = "Microsoft Office Excel";
    public string   AppVersion     { get; set; } = "16.0300";

    // ── Internal state ─────────────────────────────────────────────────────────

    private readonly XlsxWorkbook                  _workbook        = new();
    private readonly XlsxErrorLog                       _errorLog        = new();
    private readonly Dictionary<DataType, string>   _formatOverrides = [];
    private bool   _addErrorsWorksheet;
    private string _errorSheetName = "Errors";

    // ── XlsxWorkbook API ───────────────────────────────────────────────────────────

    /// <summary>Adds a worksheet and returns it for population.</summary>
    public XlsxWorksheet AddWorksheet(string name)
    {
        bool isFirst = _workbook.Sheets.Count == 0;
        var sheet = new XlsxWorksheet(
            name,
            sheetId:         _workbook.Sheets.Count + 1,
            styles:          _workbook.XlsxStyles,
            errorLog:        _errorLog,
            formatOverrides: _formatOverrides,
            tabSelected:     isFirst);
        _workbook.Sheets.Add(sheet);
        return sheet;
    }

    /// <summary>
    /// Overrides the default Excel format code for a built-in <see cref="DataType"/>.
    /// Call before adding data for the change to apply.
    /// </summary>
    public void Format(DataType dataType, string formatCode)
        => _formatOverrides[dataType] = formatCode;

    /// <summary>
    /// Enables automatic creation of an "Errors" worksheet (red tab) on
    /// <see cref="Save"/> / <see cref="GetBytes"/> when any cell errors occurred.
    /// </summary>
    public void AddErrorsWorksheet(string sheetName = "Errors")
    {
        _addErrorsWorksheet = true;
        _errorSheetName     = sheetName;
    }

    /// <summary>
    /// Returns all captured cell errors as a formatted multi-line string.
    /// Returns an empty string when no errors occurred.
    /// </summary>
    public string GetErrors()
    {
        var records = _errorLog.Records;
        if (records.Count == 0) return string.Empty;

        var sb = new StringBuilder();
        foreach (var e in records)
        {
            sb.Append("Sheet: ").Append(e.SheetName)
              .Append("  XlsxCell: ").Append(e.CellRef ?? "(null)")
              .Append("  Value: ").Append(e.AttemptedValue)
              .Append("  ").Append(e.ExceptionType).Append(": ")
              .AppendLine(e.ErrorMessage);
        }
        return sb.ToString();
    }

    // ── Save / GetBytes ────────────────────────────────────────────────────────

    public byte[] GetBytes()
    {
        using var ms = new MemoryStream();
        Save(ms);
        return ms.ToArray();
    }

    public async Task<byte[]> GetBytesAsync()
    {
        using var ms = new MemoryStream();
        await SaveAsync(ms);
        return ms.ToArray();
    }

    public void Save(string filePath)
    {
        using var stream = File.Create(filePath);
        Save(stream);
    }

    public async Task SaveAsync(string filePath)
    {
        using var stream = File.Create(filePath);
        await SaveAsync(stream);
    }

    public void Save(Stream stream)
    {
        PrepareForSave();
        WriteZip(stream);
    }

    public async Task SaveAsync(Stream stream)
    {
        PrepareForSave();
        // ZipArchive itself is synchronous; run on thread pool to free calling thread
        await Task.Run(() => WriteZip(stream));
    }

    // ── Private save helpers ───────────────────────────────────────────────────

    private void PrepareForSave()
    {
        if (_workbook.Sheets.Count == 0)
            AddWorksheet("Sheet1");

        if (_addErrorsWorksheet && _errorLog.Records.Count > 0)
        {
            try { BuildErrorSheet(); }
            catch { /* never let error-sheet building abort the save */ }
        }
    }

    private void BuildErrorSheet()
    {
        var errSheet = _workbook.Sheets.Find(s => s.Name == _errorSheetName);
        if (errSheet is null)
        {
            errSheet = new XlsxWorksheet(_errorSheetName,
                sheetId:         _workbook.Sheets.Count + 1,
                styles:          _workbook.XlsxStyles,
                errorLog:        _errorLog,
                formatOverrides: _formatOverrides,
                tabSelected:     false);
            _workbook.Sheets.Add(errSheet);
        }
        errSheet.TabColor = "FFFF0000";
        errSheet.Rows.Clear();

        // Styled header row — bold white text on dark crimson background
        var hdrOpts = new CellOptions
        {
            DataType        = DataType.String,
            Bold            = true,
            FontColor       = "FFFFFFFF",
            BackgroundColor = "FFC62828",
        };
        // Header
        int row = 1;
        errSheet.AddRow()
            .AddCell("Sheet",           hdrOpts)
            .AddCell("Cell",            hdrOpts)
            .AddCell("Attempted Value", hdrOpts)
            .AddCell("Exception Type",  hdrOpts)
            .AddCell("Error Message",   hdrOpts);

        errSheet.ColumnWidths("14", "8", "22", "22", "48");

        // Data rows
        foreach (var e in _errorLog.Records)
        {
            // Use GetOrAddRow directly to place at exact row index without incrementing next-row
            var rowModel = errSheet.GetOrAddRow(++row);
            void AddErrCell(int col, string? text)
            {
                var cellRef = CellReference.FromRowCol(row, col);
                var cell    = new XlsxCell
                {
                    Reference = cellRef,
                    Value     = CellValue.Of(text ?? string.Empty),
                };
                rowModel.Cells.Add(cell);
            }
            AddErrCell(1, e.SheetName ?? "(unknown)");
            AddErrCell(2, e.CellRef   ?? "(null)");
            AddErrCell(3, e.AttemptedValue?.ToString());
            AddErrCell(4, e.ExceptionType);
            AddErrCell(5, e.ErrorMessage);
        }
    }

    private void WriteZip(Stream stream)
    {
        var core = new XlsxCoreProperties
        {
            Creator        = Creator,
            LastModifiedBy = LastModifiedBy,
            Created        = Created,
            Modified       = Modified,
        };
        var app = new XlsxAppProperties
        {
            Application = Application,
            AppVersion  = AppVersion,
            Company     = Company,
        };

        // ── Assign global IDs ──────────────────────────────────────────────────────

        // Tables: global tableId + per-sheet localRId
        int tableId = 1;
        foreach (var sheet in _workbook.Sheets)
        {
            int localRId = 1;
            foreach (var t in sheet.Tables)
            {
                t.TableId  = tableId++;
                t.LocalRId = localRId++;
            }
        }

        // Charts / drawings / images: assign global IDs; drawing rId is after table rIds
        int drawingId = 1;
        int chartId   = 1;
        int mediaId   = 1;
        foreach (var sheet in _workbook.Sheets)
        {
            bool hasCharts = sheet.Charts.Count > 0;
            bool hasImages = sheet.Images.Count > 0;
            if (!hasCharts && !hasImages) continue;

            sheet.DrawingId       = drawingId++;
            sheet.DrawingLocalRId = sheet.Tables.Count + 1;

            foreach (var c in sheet.Charts)
            {
                c.ChartId   = chartId++;
                c.DrawingId = sheet.DrawingId;
            }

            // Image rIds follow chart rIds within the same drawing
            for (int i = 0; i < sheet.Images.Count; i++)
            {
                var img = sheet.Images[i];
                img.MediaId    = mediaId++;
                img.DrawingRId = sheet.Charts.Count + i + 1;
            }
        }

        // ── Collect IDs for content-types ─────────────────────────────────────────

        var allTableIds   = _workbook.Sheets.SelectMany(s => s.Tables).Select(t => t.TableId).ToList();
        var allDrawingIds = _workbook.Sheets.Where(s => s.DrawingId > 0).Select(s => s.DrawingId).ToList();
        var allChartIds   = _workbook.Sheets.SelectMany(s => s.Charts).Select(c => c.ChartId).ToList();
        var allImageExts  = _workbook.Sheets.SelectMany(s => s.Images).Select(img => img.Extension)
                                     .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        // ── Write ZIP ─────────────────────────────────────────────────────────────

        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        TryWrite(zip, "[Content_Types].xml",
              XlsxContentTypesPart.Generate(
                  _workbook.Sheets.Count,
                  allTableIds.Count   > 0 ? allTableIds   : null,
                  allDrawingIds.Count > 0 ? allDrawingIds : null,
                  allChartIds.Count   > 0 ? allChartIds   : null,
                  allImageExts.Count  > 0 ? allImageExts  : null));

        TryWrite(zip, "_rels/.rels",                XlsxRootRelsPart.Generate());
        TryWrite(zip, "xl/workbook.xml",            XlsxWorkbookPart.Generate(_workbook));
        TryWrite(zip, "xl/_rels/workbook.xml.rels", XlsxWorkbookRelsPart.Generate(_workbook));

        for (int i = 0; i < _workbook.Sheets.Count; i++)
        {
            var sheet = _workbook.Sheets[i];

            TryWrite(zip, $"xl/worksheets/sheet{i + 1}.xml", XlsxWorksheetPart.Generate(sheet));

            bool hasRels = sheet.Tables.Count > 0 || sheet.Charts.Count > 0 || sheet.Images.Count > 0;
            if (hasRels)
            {
                TryWrite(zip, $"xl/worksheets/_rels/sheet{i + 1}.xml.rels",
                    XlsxWorksheetRelsPart.Generate(sheet.Tables, sheet.DrawingId, sheet.DrawingLocalRId));
            }

            foreach (var t in sheet.Tables)
                TryWrite(zip, $"xl/tables/table{t.TableId}.xml", XlsxTablePart.Generate(t));

            if (sheet.Charts.Count > 0 || sheet.Images.Count > 0)
            {
                TryWrite(zip, $"xl/drawings/drawing{sheet.DrawingId}.xml",
                    XlsxDrawingPart.Generate(sheet.Charts, sheet.Images));
                TryWrite(zip, $"xl/drawings/_rels/drawing{sheet.DrawingId}.xml.rels",
                    XlsxDrawingRelsPart.Generate(sheet.Charts, sheet.Images));
                foreach (var c in sheet.Charts)
                    TryWrite(zip, $"xl/charts/chart{c.ChartId}.xml",
                        XlsxChartPart.Generate(c.Definition));
                foreach (var img in sheet.Images)
                    TryWrite(zip, $"xl/media/image{img.MediaId}.{img.Extension}", img.ImageBytes);
            }
        }

        TryWrite(zip, "xl/theme/theme1.xml", XlsxThemePart.Generate());
        TryWrite(zip, "xl/styles.xml",       XlsxStylesPart.Generate(_workbook.XlsxStyles));
        TryWrite(zip, "docProps/core.xml",   XlsxCorePropertiesPart.Generate(core));
        TryWrite(zip, "docProps/app.xml",    XlsxAppPropertiesPart.Generate(app, _workbook));
    }

    /// <summary>Writes a part to the ZIP; silently skips on exception so other parts still write.</summary>
    private static void TryWrite(ZipArchive zip, string path, byte[] content)
    {
        try { Write(zip, path, content); }
        catch { /* part generation or write failed — skip, let other parts continue */ }
    }

    private static void Write(ZipArchive zip, string path, byte[] content)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
        using var s = entry.Open();
        s.Write(content, 0, content.Length);
    }
}
