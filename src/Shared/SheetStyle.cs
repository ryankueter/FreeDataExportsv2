namespace FreeDataExportsv2;

/// <summary>
/// Default visual style applied to an entire worksheet — background color, font family,
/// size, and weight.  Individual cells that carry their own <see cref="CellOptions"/> always
/// take precedence over the sheet default.
/// </summary>
/// <remarks>
/// <para>
/// In Excel Open XML, the sheet default is expressed as a column-level style on every column
/// (including those with no data) so that empty cells also show the background color.
/// </para>
/// <para>
/// Assign via <see cref="XlsxWorksheet.SheetStyle"/> or the fluent
/// <see cref="XlsxWorksheet.ApplySheetStyle(SheetStyle)"/> method.
/// </para>
/// </remarks>
public sealed class SheetStyle
{
    /// <summary>
    /// ARGB hex background color applied to the entire sheet, e.g.
    /// <c>"FFF2F2F2"</c> (light gray).  <c>null</c> = default (white / no fill).
    /// </summary>
    public string? BackgroundColor { get; set; }

    /// <summary>
    /// XlsxFont family name, e.g. <c>"Calibri"</c>, <c>"Arial"</c>.
    /// <c>null</c> = workbook default (Calibri 11pt).
    /// </summary>
    public string? FontName { get; set; }

    /// <summary>
    /// XlsxFont size in points, e.g. <c>12</c>.
    /// <c>null</c> = workbook default (11pt).
    /// </summary>
    public double? FontSize { get; set; }

    /// <summary>
    /// ARGB hex font color, e.g. <c>"FF333333"</c> (dark gray).
    /// <c>null</c> = workbook default (black).
    /// </summary>
    public string? FontColor { get; set; }

    /// <summary>Bold text. Default <c>false</c>.</summary>
    public bool Bold { get; set; }

    /// <summary>Italic text. Default <c>false</c>.</summary>
    public bool Italic { get; set; }

    // ── Fluent builders ───────────────────────────────────────────────────────
    // Note: FontSize, FontColor, Bold, and Italic share names with their properties,
    // so those four are set via object initializer rather than fluent methods.

    /// <summary>Sets the sheet background color (ARGB hex).</summary>
    public SheetStyle Background(string argbColor) { BackgroundColor = argbColor; return this; }

    /// <summary>Sets the default font family.</summary>
    public SheetStyle XlsxFont(string fontName) { FontName = fontName; return this; }
}
