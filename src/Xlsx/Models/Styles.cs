using System.Globalization;
using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

// ── Sub-models ────────────────────────────────────────────────────────────────

internal sealed class XlsxNumFmt
{
    public int    NumFmtId   { get; set; }
    public string FormatCode { get; set; } = string.Empty;
}

internal sealed class XlsxFont
{
    public string  Name        { get; set; } = "Calibri";
    public double  Size        { get; set; } = 11;
    public bool    Bold        { get; set; }
    public bool    Italic      { get; set; }
    public bool    Underline   { get; set; }
    public bool    Strikethrough { get; set; }
    public int?    ThemeColor  { get; set; } = 1;
    public string? RgbColor    { get; set; }
    public int     Family      { get; set; } = 2;
    public string? Scheme      { get; set; } = "minor";

    public static readonly XlsxFont Default = new();
}

internal sealed class XlsxFill
{
    public string  PatternType { get; set; }
    public string? FgColor     { get; set; }
    public string? BgColor     { get; set; }

    public XlsxFill(string patternType, string? fgColor = null, string? bgColor = null)
    {
        PatternType = patternType;
        FgColor     = fgColor;
        BgColor     = bgColor;
    }
}

internal sealed class XlsxBorderSide
{
    public string? Style { get; set; }
    public string? Color { get; set; }
}

internal sealed class XlsxBorder
{
    public XlsxBorderSide Left     { get; set; } = new();
    public XlsxBorderSide Right    { get; set; } = new();
    public XlsxBorderSide Top      { get; set; } = new();
    public XlsxBorderSide Bottom   { get; set; } = new();
    public XlsxBorderSide Diagonal { get; set; } = new();
}

internal sealed class XlsxAlignment
{
    public string? Horizontal { get; set; }
    public string? Vertical   { get; set; }
    public bool    WrapText   { get; set; }
}

internal sealed class XlsxCellXf
{
    public int            NumFmtId          { get; set; }
    public int            FontId            { get; set; }
    public int            FillId            { get; set; }
    public int            BorderId          { get; set; }
    public int?           XfId             { get; set; }
    public bool           ApplyNumberFormat { get; set; }
    public bool           ApplyFont         { get; set; }
    public bool           ApplyFill         { get; set; }
    public bool           ApplyBorder       { get; set; }
    public bool           ApplyAlignment    { get; set; }
    public XlsxAlignment? XlsxAlignment        { get; set; }
}

internal sealed class XlsxCellStyle
{
    public string Name      { get; set; } = string.Empty;
    public int    XfId      { get; set; }
    public int    BuiltinId { get; set; }

    public static readonly XlsxCellStyle Normal = new() { Name = "Normal", XfId = 0, BuiltinId = 0 };
}

// ── Main XlsxStyles ──────────────────────────────────────────────────────────

internal sealed class XlsxStyles
{
    public List<XlsxNumFmt>    NumFmts      { get; } = [];
    public List<XlsxFont>      Fonts        { get; } = [XlsxFont.Default];
    public List<XlsxFill>      Fills        { get; } = [new("none"), new("gray125")];
    public List<XlsxBorder>    Borders      { get; } = [new XlsxBorder()];
    public List<XlsxCellXf>    CellStyleXfs { get; } = [new XlsxCellXf()];
    public List<XlsxCellXf>    CellXfs      { get; } = [new XlsxCellXf { XfId = 0 }];
    public List<XlsxCellStyle> CellStyles   { get; } = [XlsxCellStyle.Normal];

    // ── Format-code-only lookup (simple DataType → format mapping) ─────────────

    public int GetOrAddCellXf(string formatCode, int fontId = 0, int fillId = 0, int borderId = 0)
    {
        int numFmtId = GetOrAddNumFmt(formatCode);
        int idx = CellXfs.FindIndex(x =>
            x.NumFmtId == numFmtId && x.FontId == fontId &&
            x.FillId   == fillId   && x.BorderId == borderId &&
            x.XlsxAlignment is null);
        if (idx >= 0) return idx;

        CellXfs.Add(new XlsxCellXf
        {
            NumFmtId          = numFmtId,
            FontId            = fontId,
            FillId            = fillId,
            BorderId          = borderId,
            XfId              = 0,
            ApplyNumberFormat = numFmtId != 0,
        });
        return CellXfs.Count - 1;
    }

    // ── Full CellOptions lookup (font + fill + alignment + format) ────────────

    public int GetOrAddCellXfFromOptions(CellOptions options,
                                          Dictionary<DataType, string>? formatOverrides)
    {
        int fontId    = GetOrAddFont(options);
        int fillId    = GetOrAddFill(options);
        int borderId  = GetOrAddBorderFromOptions(options);
        string fmt    = XlsxDataTypeFormats.GetFormatCode(options.DataType, formatOverrides);
        int numFmtId  = GetOrAddNumFmt(fmt);
        var alignment = BuildAlignment(options);

        int idx = CellXfs.FindIndex(x =>
            x.NumFmtId == numFmtId && x.FontId == fontId &&
            x.FillId   == fillId   && x.BorderId == borderId &&
            AlignmentsEqual(x.XlsxAlignment, alignment));
        if (idx >= 0) return idx;

        CellXfs.Add(new XlsxCellXf
        {
            NumFmtId          = numFmtId,
            FontId            = fontId,
            FillId            = fillId,
            BorderId          = borderId,
            XfId              = 0,
            ApplyNumberFormat = numFmtId  != 0,
            ApplyFont         = fontId    != 0,
            ApplyFill         = fillId    > 1,
            ApplyBorder       = borderId  != 0,
            ApplyAlignment    = alignment is not null,
            XlsxAlignment     = alignment,
        });
        return CellXfs.Count - 1;
    }

    // ── Red-border XF used for error cells ────────────────────────────────────

    public int GetOrAddRedBorderXf()
    {
        int borderId = GetOrAddRedBorder();
        int idx = CellXfs.FindIndex(x =>
            x.NumFmtId == 0 && x.FontId == 0 &&
            x.FillId == 0 && x.BorderId == borderId);
        if (idx >= 0) return idx;

        CellXfs.Add(new XlsxCellXf
        {
            BorderId    = borderId,
            XfId        = 0,
            ApplyBorder = true,
        });
        return CellXfs.Count - 1;
    }

    // ── Private helpers ────────────────────────────────────────────────────────

    private int GetOrAddFont(CellOptions o)
    {
        bool hasOverride =
            o.FontName is not null || o.FontSize.HasValue || o.FontColor is not null ||
            o.Bold || o.Italic || o.Underline || o.Strikethrough;
        if (!hasOverride) return 0;

        var def = XlsxFont.Default;
        var fm  = new XlsxFont
        {
            Name          = o.FontName     ?? def.Name,
            Size          = o.FontSize     ?? def.Size,
            Bold          = o.Bold,
            Italic        = o.Italic,
            Underline     = o.Underline,
            Strikethrough = o.Strikethrough,
            Family        = def.Family,
            Scheme        = o.FontName is null ? def.Scheme : null,
        };

        if (o.FontColor is not null)
        {
            fm.ThemeColor = null;
            fm.RgbColor   = o.FontColor;
        }
        else
        {
            fm.ThemeColor = def.ThemeColor;
            fm.RgbColor   = null;
        }

        int idx = Fonts.FindIndex(f =>
            f.Name == fm.Name && Math.Abs(f.Size - fm.Size) < 0.001 &&
            f.Bold == fm.Bold && f.Italic == fm.Italic &&
            f.Underline == fm.Underline && f.Strikethrough == fm.Strikethrough &&
            f.ThemeColor == fm.ThemeColor && f.RgbColor == fm.RgbColor &&
            f.Scheme == fm.Scheme);
        if (idx >= 0) return idx;

        Fonts.Add(fm);
        return Fonts.Count - 1;
    }

    private int GetOrAddFill(CellOptions o)
    {
        if (o.BackgroundColor is null) return 0;
        var target = o.BackgroundColor;
        int idx = Fills.FindIndex(f =>
            f.PatternType == "solid" && f.FgColor == target);
        if (idx >= 0) return idx;

        Fills.Add(new XlsxFill("solid", fgColor: target, bgColor: "FF000000"));
        return Fills.Count - 1;
    }

    private int GetOrAddBorderFromOptions(CellOptions o)
    {
        bool hasBorder =
            o.BorderLeftStyle   != null || o.BorderRightStyle  != null ||
            o.BorderTopStyle    != null || o.BorderBottomStyle != null;
        if (!hasBorder) return 0;

        static XlsxBorderSide Side(string? style, string? color) => new()
        {
            Style = style,
            Color = style != null ? (color ?? "FF000000") : null,
        };

        var left   = Side(o.BorderLeftStyle,   o.BorderLeftColor);
        var right  = Side(o.BorderRightStyle,  o.BorderRightColor);
        var top    = Side(o.BorderTopStyle,    o.BorderTopColor);
        var bottom = Side(o.BorderBottomStyle, o.BorderBottomColor);

        int idx = Borders.FindIndex(b =>
            b.Left.Style   == left.Style   && b.Left.Color   == left.Color   &&
            b.Right.Style  == right.Style  && b.Right.Color  == right.Color  &&
            b.Top.Style    == top.Style    && b.Top.Color    == top.Color    &&
            b.Bottom.Style == bottom.Style && b.Bottom.Color == bottom.Color);
        if (idx >= 0) return idx;

        Borders.Add(new XlsxBorder { Left = left, Right = right, Top = top, Bottom = bottom });
        return Borders.Count - 1;
    }

    private int GetOrAddRedBorder()
    {
        const string red = "FFFF0000";
        int idx = Borders.FindIndex(b =>
            b.Left.Style == "thin" && b.Left.Color == red);
        if (idx >= 0) return idx;

        var side = new XlsxBorderSide { Style = "thin", Color = red };
        Borders.Add(new XlsxBorder { Left = side, Right = side, Top = side, Bottom = side });
        return Borders.Count - 1;
    }

    internal int GetOrAddNumFmt(string formatCode)
    {
        if (XlsxDataTypeFormats.BuiltInIds.TryGetValue(formatCode, out int builtIn))
            return builtIn;
        int idx = NumFmts.FindIndex(f => f.FormatCode == formatCode);
        if (idx >= 0) return NumFmts[idx].NumFmtId;
        int newId = 164 + NumFmts.Count;
        NumFmts.Add(new XlsxNumFmt { NumFmtId = newId, FormatCode = formatCode });
        return newId;
    }

    private static XlsxAlignment? BuildAlignment(CellOptions o)
    {
        if (o.HorizontalAlign is null && o.VerticalAlign is null && !o.WrapText)
            return null;
        return new XlsxAlignment
        {
            Horizontal = o.HorizontalAlign,
            Vertical   = o.VerticalAlign,
            WrapText   = o.WrapText,
        };
    }

    private static bool AlignmentsEqual(XlsxAlignment? a, XlsxAlignment? b)
    {
        if (a is null && b is null) return true;
        if (a is null || b is null) return false;
        return a.Horizontal == b.Horizontal &&
               a.Vertical   == b.Vertical   &&
               a.WrapText   == b.WrapText;
    }
}
