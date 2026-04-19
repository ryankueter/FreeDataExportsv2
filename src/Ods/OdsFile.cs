using System.IO.Compression;
using System.Text;
using FreeDataExportsv2.Internal;

namespace FreeDataExportsv2;

/// <summary>
/// Main entry point for creating an ODS (OpenDocument Spreadsheet) file.
/// API mirrors <see cref="XlsxFile"/> exactly.
/// </summary>
public sealed class OdsFile
{
    // ── Core properties ────────────────────────────────────────────────────────

    public string   Creator        { get; set; } = string.Empty;
    public string   LastModifiedBy { get; set; } = string.Empty;
    public DateTime Created        { get; set; } = DateTime.UtcNow;
    public DateTime Modified       { get; set; } = DateTime.UtcNow;
    public string   Company        { get; set; } = string.Empty;

    // ── Internal state ─────────────────────────────────────────────────────────

    private readonly XlsxWorkbook                _workbook        = new();
    private readonly XlsxErrorLog                     _errorLog        = new();
    private readonly Dictionary<DataType, string> _formatOverrides = [];
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

    /// <summary>Overrides the default format code for a <see cref="DataType"/>.</summary>
    public void Format(DataType dataType, string formatCode)
        => _formatOverrides[dataType] = formatCode;

    /// <summary>
    /// Enables automatic creation of an "Errors" worksheet on save when any cell errors occurred.
    /// </summary>
    public void AddErrorsWorksheet(string sheetName = "Errors")
    {
        _addErrorsWorksheet = true;
        _errorSheetName     = sheetName;
    }

    /// <summary>Returns all captured cell errors as a formatted multi-line string.</summary>
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
            catch { /* never abort save because of error-sheet building */ }
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

        errSheet.AddRow()
            .AddCell("Sheet",           hdrOpts)
            .AddCell("Cell",            hdrOpts)
            .AddCell("Attempted Value", hdrOpts)
            .AddCell("Exception Type",  hdrOpts)
            .AddCell("Error Message",   hdrOpts);

        errSheet.ColumnWidths("14", "8", "22", "22", "48");

        int row = 1;
        foreach (var e in _errorLog.Records)
        {
            var rowModel = errSheet.GetOrAddRow(++row);
            void AddErrCell(int col, string? text)
            {
                var cellRef = CellReference.FromRowCol(row, col);
                rowModel.Cells.Add(new XlsxCell
                {
                    Reference = cellRef,
                    Value     = CellValue.Of(text ?? string.Empty),
                });
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
        // Assign drawing / chart / image IDs
        int chartObjId = 1;
        int mediaId    = 1;
        foreach (var sheet in _workbook.Sheets)
        {
            if (sheet.Charts.Count > 0)
            {
                foreach (var c in sheet.Charts)
                    c.ChartId = chartObjId++;
            }
            foreach (var img in sheet.Images)
                img.MediaId = mediaId++;
        }

        bool hasCharts = _workbook.Sheets.Any(s => s.Charts.Count > 0);
        bool hasImages = _workbook.Sheets.Any(s => s.Images.Count > 0);

        // ODS requires mimetype to be the FIRST entry, stored (not compressed)
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        TryWrite(zip, "mimetype",
            Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.spreadsheet"),
            CompressionLevel.NoCompression);

        TryWrite(zip, "META-INF/manifest.xml",
            OdsManifestPart.Generate(_workbook, hasImages));

        TryWrite(zip, "meta.xml",
            OdsMetaPart.Generate(this));

        TryWrite(zip, "settings.xml",
            OdsSettingsPart.Generate(_workbook));

        TryWrite(zip, "styles.xml",
            OdsStylesPart.Generate());

        TryWrite(zip, "content.xml",
            OdsContentPart.Generate(_workbook, _formatOverrides));

        // Pictures (images)
        foreach (var sheet in _workbook.Sheets)
        {
            foreach (var img in sheet.Images)
            {
                TryWrite(zip, $"Pictures/image{img.MediaId}.{img.Extension}",
                    img.ImageBytes);
            }
        }

        // Chart objects (Object N/)
        foreach (var sheet in _workbook.Sheets)
        {
            foreach (var c in sheet.Charts)
            {
                TryWrite(zip, $"Object {c.ChartId}/content.xml",
                    OdsChartPart.GenerateContent(c.Definition, sheet.Name));
                TryWrite(zip, $"Object {c.ChartId}/styles.xml",
                    OdsChartPart.GenerateStyles());
                TryWrite(zip, $"Object {c.ChartId}/meta.xml",
                    OdsChartPart.GenerateMeta());
            }
        }
    }

    private static void TryWrite(ZipArchive zip, string path, byte[] content,
        CompressionLevel level = CompressionLevel.Optimal)
    {
        try
        {
            var entry = zip.CreateEntry(path, level);
            using var s = entry.Open();
            s.Write(content, 0, content.Length);
        }
        catch { /* skip failed parts so others still write */ }
    }
}
