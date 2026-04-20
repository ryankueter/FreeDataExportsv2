namespace FreeDataExportsv2;

/// <summary>
/// Extensible cell-formatting options. Pass to
/// <see cref="XlsxRowBuilder.AddCell(object?,CellOptions)"/> for fine-grained control.
/// </summary>
public sealed class CellOptions
{
    public DataType DataType { get; set; } = DataType.General;

    // XlsxFont
    public string? FontName      { get; set; }
    public double? FontSize      { get; set; }
    /// <summary>ARGB hex, e.g. "FF000000".</summary>
    public string? FontColor     { get; set; }
    public bool    Bold          { get; set; }
    public bool    Italic        { get; set; }
    public bool    Underline     { get; set; }
    public bool    Strikethrough { get; set; }

    // XlsxFill
    /// <summary>ARGB hex background color, e.g. "FFFFFF00" (yellow).</summary>
    public string? BackgroundColor { get; set; }

    // XlsxAlignment
    /// <summary>Use <see cref="FreeDataExportsv2.HorizontalAlign"/> constants: Left, Center, Right, Fill, Justify, General.</summary>
    public string? HorizontalAlign { get; set; }
    /// <summary>Use <see cref="FreeDataExportsv2.VerticalAlign"/> constants: Top, Center, Bottom, Justify.</summary>
    public string? VerticalAlign   { get; set; }
    public bool    WrapText        { get; set; }

    // XlsxBorder — omit a side to leave it unstyled
    /// <summary>Use <see cref="BorderStyle"/> constants: Thin, Medium, Thick, Dashed, Dotted, Double, Hair, MediumDashed, DashDot, MediumDashDot, DashDotDot, MediumDashDotDot, SlantDashDot.</summary>
    public string? BorderLeftStyle   { get; set; }
    /// <summary>ARGB hex, e.g. "FF000000". Defaults to black when omitted.</summary>
    public string? BorderLeftColor   { get; set; }
    /// <inheritdoc cref="BorderLeftStyle"/>
    public string? BorderRightStyle  { get; set; }
    /// <inheritdoc cref="BorderLeftColor"/>
    public string? BorderRightColor  { get; set; }
    /// <inheritdoc cref="BorderLeftStyle"/>
    public string? BorderTopStyle    { get; set; }
    /// <inheritdoc cref="BorderLeftColor"/>
    public string? BorderTopColor    { get; set; }
    /// <inheritdoc cref="BorderLeftStyle"/>
    public string? BorderBottomStyle { get; set; }
    /// <inheritdoc cref="BorderLeftColor"/>
    public string? BorderBottomColor { get; set; }
}
