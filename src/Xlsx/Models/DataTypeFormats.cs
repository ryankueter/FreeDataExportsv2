using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

internal static class XlsxDataTypeFormats
{
    internal static readonly Dictionary<DataType, string> Defaults = new()
    {
        [DataType.General]      = "General",
        [DataType.Number]       = "General",
        [DataType.Integer]      = "0",
        [DataType.Currency]     = "\"$\"#,##0.00",
        [DataType.Accounting]   = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
        [DataType.Percentage]   = "0.00%",
        [DataType.WholePercent] = "0%",
        [DataType.Fraction]     = "# ??/??",
        [DataType.Scientific]   = "0.00E+00",
        [DataType.ShortDate]    = "m/d/yyyy",
        [DataType.LongDate]     = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy",
        [DataType.DateTime]     = "m/d/yy h:mm",
        [DataType.DateTime24]   = "m/d/yy h:mm",
        [DataType.Time12h]      = "h:mm:ss AM/PM",
        [DataType.Time24h]      = "h:mm:ss",
        [DataType.PhoneUS]      = "(###) ###-####",
        [DataType.Text]         = "@",
        [DataType.Zip]          = "00000",
        [DataType.Thousands]    = "#,##0",
        [DataType.Thousands2]   = "#,##0.00",
    };

    // Built-in Excel numFmtIds — these must NOT be added to <numFmts> section
    internal static readonly Dictionary<string, int> BuiltInIds = new(StringComparer.Ordinal)
    {
        ["General"]               = 0,
        ["0"]                     = 1,
        ["0.00"]                  = 2,
        ["#,##0"]                 = 3,
        ["#,##0.00"]              = 4,
        ["0%"]                    = 9,
        ["0.00%"]                 = 10,
        ["0.00E+00"]              = 11,
        ["# ??/??"]               = 12,
        ["# ???/???"]             = 13,
        ["m/d/yyyy"]              = 14,
        ["d-mmm-yy"]              = 15,
        ["d-mmm"]                 = 16,
        ["mmm-yy"]                = 17,
        ["h:mm AM/PM"]            = 18,
        ["h:mm:ss AM/PM"]         = 19,
        ["h:mm"]                  = 20,
        ["h:mm:ss"]               = 21,
        ["m/d/yy h:mm"]           = 22,
        ["#,##0 ;(#,##0)"]        = 37,
        ["#,##0 ;[Red](#,##0)"]   = 38,
        ["#,##0.00;(#,##0.00)"]   = 39,
        ["#,##0.00;[Red](#,##0.00)"] = 40,
        ["mm:ss"]                 = 45,
        ["[h]:mm:ss"]             = 46,
        ["mmss.0"]                = 47,
        ["##0.0E+0"]              = 48,
        ["@"]                     = 49,
    };

    internal static bool IsDateType(DataType dt) => dt is
        DataType.ShortDate or DataType.LongDate or
        DataType.DateTime  or DataType.DateTime24 or
        DataType.Time12h   or DataType.Time24h;

    internal static string GetFormatCode(DataType dt, Dictionary<DataType, string>? overrides)
    {
        if (overrides is not null && overrides.TryGetValue(dt, out var ov)) return ov;
        return Defaults.TryGetValue(dt, out var def) ? def : "General";
    }
}
