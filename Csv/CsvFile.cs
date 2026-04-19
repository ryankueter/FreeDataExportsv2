using System.Globalization;
using System.Text;

namespace FreeDataExportsv2;

/// <summary>
/// Produces a single CSV document.  Rows are added directly on the package —
/// no worksheet abstraction is needed.
/// </summary>
/// <example>
/// <code>
/// var csv = new CsvFile();
/// csv.AddRow("OrderId", "Item", "Units", "Price", "OrderDate", "SalesAssoc", "Delivered");
/// foreach (var o in orders)
///     csv.AddRow(o.OrderId, o.Item, o.Units, o.Price, o.OrderDate, o.SalesAssociate, o.Delivered);
/// await csv.SaveAsync(path);
/// </code>
/// </example>
public sealed class CsvFile
{
    private readonly List<List<string>> _rows = [];

    /// <summary>Field delimiter. Default is comma.</summary>
    public string Delimiter  { get; set; } = ",";

    /// <summary>Prepend a UTF-8 BOM for Excel compatibility. Default is true.</summary>
    public bool   IncludeBom { get; set; } = true;

    /// <summary>Line ending sequence. Default is CR+LF (RFC 4180).</summary>
    public string LineEnding { get; set; } = "\r\n";

    // ── XlsxRow population ─────────────────────────────────────────────────────────

    /// <summary>
    /// Appends a row whose cells are the formatted string representations of
    /// <paramref name="values"/>. Returns <c>this</c> for chaining.
    /// </summary>
    /// <remarks>
    /// Auto-formatting rules (no <see cref="DataType"/> required):
    /// <list type="bullet">
    ///   <item><c>null</c> → empty string</item>
    ///   <item><see cref="bool"/> → <c>TRUE</c> / <c>FALSE</c></item>
    ///   <item><see cref="DateTime"/> → <c>M/d/yyyy</c> (invariant culture)</item>
    ///   <item><see cref="DateTimeOffset"/> → date portion, <c>M/d/yyyy</c></item>
    ///   <item>Numeric types → invariant-culture string</item>
    ///   <item>Everything else → <c>ToString()</c></item>
    /// </list>
    /// </remarks>
    public CsvFile AddRow(params object?[] values)
    {
        var row = new List<string>(values.Length);
        foreach (var v in values)
            row.Add(FormatValue(v));
        _rows.Add(row);
        return this;
    }

    // ── Output ─────────────────────────────────────────────────────────────────

    /// <summary>Serializes all rows to bytes (with optional BOM).</summary>
    public byte[] GetBytes()
    {
        using var ms = new MemoryStream();
        Save(ms);
        return ms.ToArray();
    }

    /// <summary>Serializes all rows to bytes asynchronously.</summary>
    public async Task<byte[]> GetBytesAsync()
    {
        using var ms = new MemoryStream();
        await SaveAsync(ms);
        return ms.ToArray();
    }

    /// <summary>Saves all rows to a file.</summary>
    public void Save(string path)
    {
        using var stream = File.Create(path);
        Save(stream);
    }

    /// <summary>Saves all rows to a stream.</summary>
    public void Save(Stream stream)
    {
        var text = Render();
        WriteToStream(stream, text, IncludeBom);
    }

    /// <summary>Saves all rows to a file asynchronously.</summary>
    public async Task SaveAsync(string path)
    {
        using var stream = File.Create(path);
        await SaveAsync(stream);
    }

    /// <summary>Saves all rows to a stream asynchronously.</summary>
    public async Task SaveAsync(Stream stream)
    {
        var text = Render();
        await Task.Run(() => WriteToStream(stream, text, IncludeBom));
    }

    // ── Private helpers ────────────────────────────────────────────────────────

    private string Render()
    {
        if (_rows.Count == 0) return string.Empty;

        char delimChar = Delimiter.Length == 1 ? Delimiter[0] : ',';
        var  sb        = new StringBuilder();

        foreach (var row in _rows)
        {
            for (int c = 0; c < row.Count; c++)
            {
                if (c > 0) sb.Append(Delimiter);
                sb.Append(CsvQuote(row[c], delimChar));
            }
            sb.Append(LineEnding);
        }

        return sb.ToString();
    }

    private static string FormatValue(object? value) => value switch
    {
        null              => string.Empty,
        bool   b          => b ? "TRUE" : "FALSE",
        DateTime dt       => dt.ToString("M/d/yyyy", CultureInfo.InvariantCulture),
        DateTimeOffset dto => dto.ToString("M/d/yyyy", CultureInfo.InvariantCulture),
        decimal d         => d.ToString(CultureInfo.InvariantCulture),
        double  d         => d.ToString(CultureInfo.InvariantCulture),
        float   f         => f.ToString(CultureInfo.InvariantCulture),
        _                 => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty,
    };

    /// <summary>
    /// Wraps <paramref name="text"/> in double-quotes if it contains the delimiter,
    /// a double-quote, CR, or LF (RFC 4180). Internal quotes are doubled.
    /// </summary>
    private static string CsvQuote(string text, char delimChar)
    {
        bool needs = text.Contains(delimChar)
                  || text.Contains('"')
                  || text.Contains('\r')
                  || text.Contains('\n');

        return needs ? "\"" + text.Replace("\"", "\"\"") + "\"" : text;
    }

    private static void WriteToStream(Stream stream, string text, bool bom)
    {
        var encoding = bom ? new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)
                           : new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var bytes = encoding.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }
}
