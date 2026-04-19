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
    /// <summary>"left", "center", "right", "fill", "justify", "general"</summary>
    public string? HorizontalAlign { get; set; }
    /// <summary>"top", "center", "bottom", "justify"</summary>
    public string? VerticalAlign   { get; set; }
    public bool    WrapText        { get; set; }
}
