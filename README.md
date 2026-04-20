# FreeDataExportsv2

**Author:** Ryan Kueter &nbsp;|&nbsp; **Updated:** April 2026

A lightweight, zero-dependency .NET library for generating `.xlsx`, `.ods`, and `.csv` files.
It includes charts, images, tables, and custom formatting options, and produces valid Open XML (OOXML)
and OpenDocument (ODS) files — no Excel or LibreOffice installation required.

| Target Frameworks |
|-------------------|
| .NET 8 - .NET 10 |

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [Excel Export (`XlsxFile`)](#excel-export-xlsxfile)
3. [LibreOffice Export (`OdsFile`)](#libreoffice-export-odsfile)
4. [Worksheets](#worksheets)
5. [Adding Data — `AddRow` / `AddCell`](#adding-data--addrow--addcell)
6. [Direct Cell Access — `SetCell`](#direct-cell-access--setcell)
7. [Sheet-Level Defaults (`SheetStyle`)](#sheet-level-defaults-sheetstyle)
8. [Data Types (`DataType`)](#data-types-datatype)
9. [Overriding Format Codes](#overriding-format-codes)
10. [Cell Formatting (`CellOptions`)](#cell-formatting-celloptions)
11. [Column Widths](#column-widths)
12. [Excel Tables](#excel-tables)
13. [Charts](#charts)
14. [Images](#images)
15. [CSV Export (`CsvFile`)](#csv-export-csvfile)
16. [Error Handling](#error-handling)
17. [Saving / Getting Bytes](#saving--getting-bytes)
18. [DataType Reference](#datatype-reference)
19. [CellOptions Reference](#celloptions-reference)
20. [XlsxTableStyles Reference](#xlshtablestyles-reference)
21. [License](#license)

---

## Quick Start

```
using FreeDataExportsv2;

// XLSX — Microsoft Excel / Open XML
var workbook = new XlsxFile
{
    Creator = "Summit Ridge Outfitters",
    LastModifiedBy = "Summit Ridge Outfitters",
    Created = new DateTime(2026, 3, 31),
    Modified = DateTime.Now,
    Company = "Summit Ridge Outfitters, LLC",
};

// ODS  — LibreOffice Calc / OpenDocument (identical API, different class)
// var workbook = new OdsFile { Creator = "Jane Smith" };

var sheet = workbook.AddWorksheet("Sales");
sheet.TabColor = "FF1A3A5C"; // dark navy

// Header row
sheet.AddRow()
    .AddCell("Product",  DataType.String)
    .AddCell("Units",    DataType.String)
    .AddCell("Revenue",  DataType.String);

// Data rows
sheet.AddRow()
    .AddCell("Widget A", DataType.String)
    .AddCell(142,        DataType.Number)
    .AddCell(1419.58m,   DataType.Currency);

// Cell options
var navyBg = new CellOptions
{
    DataType        = DataType.String,
    FontSize        = 18,
    Bold            = true,
    FontColor       = "FFFFFFFF",
    BackgroundColor = "FF1A3A5C",
    BorderBottomColor = "FFFFFFFF",
    BorderBottomStyle = "medium"
};
sheet.AddRow()
    .AddCell("--", navyBg)
    .AddCell("--", navyBg)
    .AddCell("--", navyBg);

sheet.ColumnWidths("22", "18", "28");

workbook.AddErrorsWorksheet();
workbook.SaveAsync("Sales.xlsx");  // or "Sales.ods" for OdsFile
```

---

## Excel Export (`XlsxFile`)

`XlsxFile` is the entry point for generating `.xlsx` files.  All worksheets, format overrides,
and error-handling options are configured on it before calling `Save`.
See [LibreOffice Export (`OdsFile`)](#libreoffice-export-odsfile) for the equivalent OpenDocument class.

```
var workbook = new XlsxFile
{
    Creator        = "Ryan Kueter",
    LastModifiedBy = "Ryan Kueter",
    Created        = new DateTime(2026, 1, 1),
    Modified       = DateTime.Now,
    Company        = "Kueter Development",
    Application    = "Microsoft Office Excel",  // default
    AppVersion     = "16.0300",                 // default
};
```

### Key methods

| Method | Description |
|---|---|
| `AddWorksheet(name)` | Creates and returns a new `XlsxWorksheet`. |
| `Format(DataType, formatCode)` | Overrides the default Excel format code for a `DataType`. |
| `AddErrorsWorksheet(name?)` | Opts in to automatic error-tab creation (see [Error Handling](#error-handling)). |
| `GetErrors()` | Returns all captured cell errors as a formatted string. |
| `Save(path)` | Synchronous save to a file path. |
| `Save(stream)` | Synchronous write to any `Stream`. |
| `SaveAsync(path)` | Asynchronous save to a file path. |
| `SaveAsync(stream)` | Asynchronous write to any `Stream`. |
| `GetBytes()` | Returns the complete `.xlsx` file as a `byte[]`. |
| `GetBytesAsync()` | Returns the complete `.xlsx` file as a `byte[]` asynchronously. |

---

## LibreOffice Export (`OdsFile`)

`OdsFile` produces `.ods` (OpenDocument Spreadsheet) files — the native format for LibreOffice
Calc and other OpenDocument-compatible applications.  Its public API is **identical** to
`XlsxFile`, so switching between the two requires only changing the class name.

```
using FreeDataExportsv2;

var workbook = new OdsFile
{
    Creator        = "Jane Smith",
    LastModifiedBy = "Jane Smith",
    Created        = new DateTime(2026, 1, 1),
    Modified       = DateTime.Now,
    Company        = "Acme Corp",
};

var sheet = workbook.AddWorksheet("Sales");
sheet.TabColor = "FF007B6E"; // teal tab

sheet.AddRow()
    .AddCell("Product",  DataType.String)
    .AddCell("Units",    DataType.String)
    .AddCell("Revenue",  DataType.String);

sheet.AddRow()
    .AddCell("Widget A", DataType.String)
    .AddCell(142,        DataType.Number)
    .AddCell(1419.58m,   DataType.Currency);

await workbook.SaveAsync("Sales.ods");
```

### Key methods

The `OdsFile` methods mirror `XlsxFile` exactly:

| Method | Description |
|---|---|
| `AddWorksheet(name)` | Creates and returns a new `XlsxWorksheet`. |
| `Format(DataType, formatCode)` | Overrides the default format code for a `DataType`. |
| `AddErrorsWorksheet(name?)` | Opts in to automatic error-tab creation. |
| `GetErrors()` | Returns all captured cell errors as a formatted string. |
| `Save(path)` | Synchronous save to a file path. |
| `Save(stream)` | Synchronous write to any `Stream`. |
| `SaveAsync(path)` | Asynchronous save to a file path. |
| `SaveAsync(stream)` | Asynchronous write to any `Stream`. |
| `GetBytes()` | Returns the complete `.ods` file as a `byte[]`. |
| `GetBytesAsync()` | Returns the complete `.ods` file as a `byte[]` asynchronously. |

### Differences from `XlsxFile`

| Feature | `XlsxFile` | `OdsFile` |
|---|---|---|
| File format | OOXML `.xlsx` | OpenDocument `.ods` |
| Worksheets, cells, styles | ✅ | ✅ |
| `SheetStyle` (sheet defaults) | ✅ | ✅ |
| Charts | ✅ | ✅ |
| Images | ✅ | ✅ |
| Error handling / errors worksheet | ✅ | ✅ |
| Excel Tables (`AddTable`) | ✅ | ❌ Not supported in ODS |
| `Application` / `AppVersion` properties | ✅ | ❌ Not applicable |

> **ODS package structure:** the `.ods` file is a ZIP archive that follows the ODS spec —
> `mimetype` is the first entry (stored, not compressed), followed by `content.xml`,
> `styles.xml`, `meta.xml`, `settings.xml`, `META-INF/manifest.xml`, embedded images
> (`Pictures/`), and chart sub-documents (`Object N/`).

---

## Worksheets

```
var orders    = workbook.AddWorksheet("Orders");
var inventory = workbook.AddWorksheet("Inventory");

// Optional — ARGB hex tab colour
orders.TabColor    = "FF00B050"; // green
inventory.TabColor = "FF0070C0"; // blue
```

The first worksheet added is automatically tab-selected (active on open).

> **Note:** If you call `Save` / `GetBytes` without adding any worksheets, a blank worksheet
> named `"Sheet1"` is inserted automatically so the file is always valid.

---

## Adding Data — `AddRow` / `AddCell`

`AddRow()` advances to the next row and returns an `XlsxRowBuilder` for fluent chaining.  
`AddCell(value, dataType)` accepts **any** .NET value and converts it automatically.

```
// Column headers
sheet.AddRow()
    .AddCell("OrderId",   DataType.String)
    .AddCell("Item",      DataType.String)
    .AddCell("Price",     DataType.String);

// Data rows
foreach (var o in orders)
{
    sheet.AddRow()
        .AddCell(o.OrderId,   DataType.Number)
        .AddCell(o.Item,      DataType.String)
        .AddCell(o.Price,     DataType.Currency);
}
```

`AddCell` has two overloads:

```
// Simple: value + DataType
.AddCell(value, DataType.Currency)

// Advanced: value + full CellOptions (font, fill, alignment, …)
.AddCell(value, new CellOptions { DataType = DataType.Currency, Bold = true })
```

### Supported value types for automatic coercion

`int`, `long`, `float`, `double`, `decimal`, `string`, `bool`, `DateTime`, `CellValue`
(the library's discriminated union — used for formulas and error cells).

---

## Direct Cell Access — `SetCell`

Write to any cell by A1-style reference or 1-based row/column indices, without calling `AddRow`.
Useful for sparse writes, templates, or when you need to target a specific location.

```
// By cell reference
sheet.SetCell("A1", "Label",      DataType.String);
sheet.SetCell("B1", 3.14,         DataType.Number);
sheet.SetCell("C1", DateTime.Now, DataType.ShortDate);

// With full CellOptions
sheet.SetCell("D1", -99.99m, new CellOptions
{
    DataType  = DataType.Currency,
    FontColor = "FFFF0000",   // red
    Bold      = true,
});

// By row/column index (1-based)
sheet.SetCell(2, 1, "Row 2, Col A", DataType.String);
sheet.SetCell(2, 2, 42,             DataType.Number);
```

`SetCell` returns `this` (the `XlsxWorksheet`) for chaining. It is safe to call in
any order — rows and columns are created as needed. Any conversion error is logged
(the cell gets a red border) rather than throwing.

---

## Sheet-Level Defaults (`SheetStyle`)

Set a default background color, font family, font size, and font color for an entire worksheet
with a single `SheetStyle` object.  Every cell that doesn't carry its own `CellOptions` inherits
the sheet default — including completely empty cells, so the background fills the visible grid.

```
var sheet = workbook.AddWorksheet("Report");

// Object-initializer style
sheet.SheetStyle = new SheetStyle
{
    BackgroundColor = "FFFFF8F0",   // warm cream (ARGB)
    FontName        = "Georgia",
    FontSize        = 12,
    FontColor       = "FF2C2C2C",   // near-black
};

// Fluent style — Background() and Font() are chainable;
// FontSize, FontColor, Bold, and Italic are set via object initializer
sheet.SheetStyle = new SheetStyle { FontSize = 12, FontColor = "FF2C2C2C" }
    .Background("FFFFF8F0")
    .Font("Georgia");

// Or assign via the worksheet's own fluent method
workbook.AddWorksheet("Light")
    .ApplySheetStyle(new SheetStyle { BackgroundColor = "FFF2F2F2" });
```

`CellOptions` on individual cells always takes precedence over the sheet default.
A header row with its own `BackgroundColor` will display its own color, not the sheet color.

### SheetStyle properties

| Property | Type | Description |
|---|---|---|
| `BackgroundColor` | `string?` | ARGB hex fill for the whole sheet, e.g. `"FFF2F2F2"` |
| `FontName` | `string?` | Default font family, e.g. `"Arial"` |
| `FontSize` | `double?` | Default point size, e.g. `12` |
| `FontColor` | `string?` | ARGB hex text color, e.g. `"FF333333"` |
| `Bold` | `bool` | Bold weight for the sheet default font |
| `Italic` | `bool` | Italic style for the sheet default font |

### SheetStyle fluent methods

| Method | Description |
|---|---|
| `Background(argb)` | Sets `BackgroundColor` — chainable |
| `Font(name)` | Sets `FontName` — chainable |

> `FontSize`, `FontColor`, `Bold`, and `Italic` share names with their properties so they
> are set directly via object-initializer syntax rather than fluent methods (see example above).

### `XlsxWorksheet.ApplySheetStyle`

```
workbook.AddWorksheet("Styled")
    .ApplySheetStyle(new SheetStyle { BackgroundColor = "FFF2F2F2", FontSize = 11 });
```

Returns `this` for chaining with other `XlsxWorksheet` calls.

---

## Data Types (`DataType`)

The `DataType` enum controls both the Excel cell type and the display format.

### Numeric

| Value | Format code | Example display |
|---|---|---|
| `General` | `General` | (auto) |
| `Number` | `General` | 3.14 |
| `Integer` | `0` | 42 |
| `Currency` | `"$"#,##0.00` | $1,234.56 |
| `Accounting` | full accounting | (1,234.56) |
| `Thousands` | `#,##0` | 1,500,000 |
| `Thousands2` | `#,##0.00` | 1,500,000.75 |
| `Percentage` | `0.00%` | 12.34% |
| `WholePercent` | `0%` | 75% |
| `Fraction` | `# ??/??` | 1/2 |
| `Scientific` | `0.00E+00` | 1.23E+04 |
| `PhoneUS` | `(###) ###-####` | (405) 555-1234 |
| `Zip` | `00000` | 07631 |
| `Text` | `@` | (forced text) |

### Date / Time

| Value | Format code | Example display |
|---|---|---|
| `ShortDate` | `m/d/yyyy` | 4/16/2026 |
| `LongDate` | `[$-F800]dddd\,\ mmmm\ dd\,\ yyyy` | Thursday, April 16, 2026 |
| `DateTime` | `m/d/yy h:mm` | 4/16/26 14:30 |
| `DateTime24` | `m/d/yy h:mm` | 4/16/26 14:30 (overridable) |
| `Time12h` | `h:mm:ss AM/PM` | 2:30:00 PM |
| `Time24h` | `h:mm:ss` | 14:30:00 |

### Other

| Value | Notes |
|---|---|
| `String` | Inline string — no format code |
| `Boolean` | Excel `TRUE`/`FALSE` |
| `Formula` | Pass formula text (without `=`); computed by Excel |
| `Error` | Pass a `CellValue.AsError(ErrorCode.X)` |

---

## Overriding Format Codes

Change the default format code for any `DataType` before writing data:

```
workbook.Format(DataType.DateTime24, @"m/d/yy\ h:mm");
workbook.Format(DataType.Currency,   "£#,##0.00");
workbook.Format(DataType.PhoneUS,    "000-000-0000");
```

Overrides apply to all worksheets.

---

## Cell Formatting (`CellOptions`)

Pass a `CellOptions` object as the second argument to `AddCell` for fine-grained control
over font, background colour, alignment, and borders.

```
// Reusable style
var headerStyle = new CellOptions
{
    DataType        = DataType.String,
    FontSize        = 14,
    FontColor       = "FFFFFFFF",   // white text (ARGB)
    Bold            = true,
    BackgroundColor = "FF203864",   // dark navy (ARGB)
};

sheet.AddRow()
    .AddCell("Name",   headerStyle)
    .AddCell("Amount", headerStyle);

// Inline style
sheet.AddRow()
    .AddCell(-99.99m, new CellOptions
    {
        DataType  = DataType.Currency,
        FontColor = "FFFF0000",   // red text
        Bold      = true,
    });
```

### Borders

Each side of the cell border is controlled independently.  Omit a side to leave it unstyled.

```
// Box border — all four sides, thin black line
var boxBorder = new CellOptions
{
    DataType         = DataType.String,
    BorderLeftStyle  = "thin",
    BorderRightStyle = "thin",
    BorderTopStyle   = "thin",
    BorderBottomStyle = "thin",
    // BorderXxxColor defaults to black ("FF000000") when omitted
};

// Bottom border only — common totals-row separator
var totalsSep = new CellOptions
{
    DataType          = DataType.Currency,
    Bold              = true,
    BorderBottomStyle = "medium",
    BorderBottomColor = "FF1A3A5C",   // navy underline
};

// Thick outer box with a custom colour
var thickBox = new CellOptions
{
    DataType          = DataType.String,
    BorderLeftStyle   = "thick",
    BorderLeftColor   = "FF007B6E",
    BorderRightStyle  = "thick",
    BorderRightColor  = "FF007B6E",
    BorderTopStyle    = "thick",
    BorderTopColor    = "FF007B6E",
    BorderBottomStyle = "thick",
    BorderBottomColor = "FF007B6E",
};

sheet.AddRow()
    .AddCell("Subtotal", totalsSep)
    .AddCell(1234.56m,   totalsSep);
```

#### Border style values

| Value | Description |
|---|---|
| `"thin"` | Thin solid line (most common) |
| `"medium"` | Medium solid line |
| `"thick"` | Thick solid line |
| `"dashed"` | Thin dashed line |
| `"mediumDashed"` | Medium dashed line |
| `"dotted"` | Dotted line |
| `"hair"` | Hairline (thinnest possible) |
| `"double"` | Double line |
| `"dashDot"` | Dash-dot pattern |
| `"mediumDashDot"` | Medium dash-dot |
| `"dashDotDot"` | Dash-dot-dot pattern |
| `"mediumDashDotDot"` | Medium dash-dot-dot |
| `"slantDashDot"` | Slanted dash-dot |

### CellOptions properties

| Property | Type | Description |
|---|---|---|
| `DataType` | `DataType` | Data type and number format (default: `General`) |
| `FontName` | `string?` | Font family name, e.g. `"Arial"` |
| `FontSize` | `double?` | Point size, e.g. `14` |
| `FontColor` | `string?` | ARGB hex, e.g. `"FFFFFFFF"` (white) |
| `Bold` | `bool` | Bold weight |
| `Italic` | `bool` | Italic style |
| `Underline` | `bool` | Single underline |
| `Strikethrough` | `bool` | Strikethrough |
| `BackgroundColor` | `string?` | ARGB hex cell background, e.g. `"FFFFFF00"` (yellow) |
| `HorizontalAlign` | `string?` | `"left"`, `"center"`, `"right"`, `"fill"`, `"justify"` |
| `VerticalAlign` | `string?` | `"top"`, `"center"`, `"bottom"`, `"justify"` |
| `WrapText` | `bool` | Wrap text within the cell |
| `BorderLeftStyle` | `string?` | Left border line style (see Border style values above) |
| `BorderLeftColor` | `string?` | Left border ARGB color (default: `"FF000000"`) |
| `BorderRightStyle` | `string?` | Right border line style |
| `BorderRightColor` | `string?` | Right border ARGB color |
| `BorderTopStyle` | `string?` | Top border line style |
| `BorderTopColor` | `string?` | Top border ARGB color |
| `BorderBottomStyle` | `string?` | Bottom border line style |
| `BorderBottomColor` | `string?` | Bottom border ARGB color |

> **ARGB format:** all colour values use the 8-character ARGB hex string `"FFRRGGBB"` where the
> first two characters are the alpha channel (always `FF` for fully opaque).

---

## Column Widths

Set column widths after adding all rows. Returns `this` for chaining.

```
orders.ColumnWidths("10", "14", "7", "12", "28", "14", "10");
inventory.ColumnWidths("10", "16", "10", "14");
```

### Units

`ColumnWidths` accepts values in **Excel character units** — the number of characters of the
maximum-digit-width font that fit in the column at the workbook's default font size. A value
of `"10"` is roughly 10 characters wide.

| Format | Internal unit | What the library does |
|---|---|---|
| `XlsxFile` | Character units | Written directly to `customWidth` in the OOXML sheet XML |
| `OdsFile` | Centimetres (ODS spec) | Converted automatically: `chars × 0.2258 cm` |

You pass the **same numeric values** to `ColumnWidths` regardless of whether you are
generating XLSX or ODS — the unit conversion for ODS happens internally. A column set
to `"14"` will render at approximately the same visual width in both Excel and LibreOffice.

---

## Excel Tables

> **`XlsxFile` only.** `AddTable` is not available on `OdsFile` — ODS has no equivalent concept.

Attach a styled, filterable Excel table to any range using `AddTable`.  
The table definition is fluent — chain calls to configure it.

```
// 1. Write your data rows first (header + data + optional blank totals row)
sheet.AddRow()
    .AddCell("Item ID",    DataType.String)
    .AddCell("Item",       DataType.String)
    .AddCell("In Stock",   DataType.String)
    .AddCell("Unit Price", DataType.String);

foreach (var item in inventory)
{
    sheet.AddRow()
        .AddCell(item.ItemId,    DataType.Number)
        .AddCell(item.Item,      DataType.String)
        .AddCell(item.Stock,     DataType.Number)
        .AddCell(item.UnitPrice, DataType.Currency);
}

sheet.AddRow(); // blank row reserved for totals

// 2. Attach the table over the full range (header row through totals row)
sheet.AddTable(
    cellRange:  "A1:D7",
    definition: new XlsxTableDefinition("InventoryTable")
        .Style(XlsxTableStyles.Medium3)
        .ShowTotalsRow()
        .AddColumn("Item ID")
        .AddColumn("Item")
        .AddColumn("In Stock",   XlsxTotalsRowFunction.Sum)
        .AddColumn("Unit Price", XlsxTotalsRowFunction.Average));
```

When `ShowTotalsRow()` is enabled, the library automatically inserts `SUBTOTAL` formula
cells in the last row for every column that has a non-`None` `XlsxTotalsRowFunction`.

### XlsxTableDefinition API

| Method | Description |
|---|---|
| `new XlsxTableDefinition(name)` | Creates a table with the given internal and display name |
| `AddColumn(name, totalsFunction?)` | Appends a column; optional totals aggregate |
| `Style(styleName)` | Sets the table style (see [XlsxTableStyles Reference](#xlshtablestyles-reference)) |
| `ShowTotalsRow(bool?)` | Enables the totals row |
| `RowStripes(bool?)` | Alternating row shading (default: true) |
| `ColStripes(bool?)` | Alternating column shading (default: false) |
| `FirstColumn(bool?)` | Highlight the first column |
| `LastColumn(bool?)` | Highlight the last column |

### Properties (also settable directly)

```
var def = new XlsxTableDefinition("MyTable")
{
    StyleName      = XlsxTableStyles.Dark2,
    HasTotalsRow   = true,
    ShowRowStripes = true,
};
def.Columns.Add(new XlsxTableColumn("Revenue", XlsxTotalsRowFunction.Sum));
```

### XlsxTotalsRowFunction values

`None`, `Sum`, `Average`, `Count`, `CountNums`, `Max`, `Min`, `StdDev`, `Var`

---

## Charts

Embed one or more charts on any worksheet with `AddChart`.  Charts reference data via
A1-style sheet references — the data can live on the same sheet or any other sheet.

```
// Column chart anchored to columns A–G, rows 8–23
sheet.AddChart(
    new ChartDefinition("Sales by Item")
        .Type(ChartType.Column)
        .Legend("b")                         // legend below chart
        .Series("Units Sold",
                   valuesRef:   "Orders!$C$2:$C$6",
                   categoryRef: "Orders!$B$2:$B$6"),
    new ObjectAnchor { FromCol = 0, FromRow = 7, ToCol = 6, ToRow = 22 });

// Pie chart, no title, legend on the right
sheet.AddChart(
    new ChartDefinition()
        .Type(ChartType.Pie)
        .Legend("r")
        .Series("Revenue",
                   valuesRef:   "Sheet1!$B$2:$B$6",
                   categoryRef: "Sheet1!$A$2:$A$6"),
    new ObjectAnchor { FromCol = 7, FromRow = 7, ToCol = 13, ToRow = 22 });
```

### ChartDefinition API

| Method / Property | Description |
|---|---|
| `new ChartDefinition(title?)` | Creates a chart definition with an optional title |
| `Title` | Property — sets the chart title (also set via constructor) |
| `Type(ChartType)` | Sets the chart type (see below) |
| `Legend(string?)` | Sets legend position: `"r"`, `"l"`, `"t"`, `"b"`, `"tr"` |
| `HideLegend()` | Hides the legend |
| `Series(name, valuesRef, categoryRef?)` | Appends a data series |

### ChartType values

| Value | Description |
|---|---|
| `Column` | Vertical bars (default) |
| `Bar` | Horizontal bars |
| `Line` | Line chart |
| `Pie` | Pie chart (no axes) |
| `Area` | Filled area chart |

### ObjectAnchor

Controls the chart's or image's position on the sheet using zero-based row/column indices
(column 0 = A, row 0 = row 1):

```
new ObjectAnchor
{
    FromCol = 0,  // top-left column (0 = A)
    FromRow = 7,  // top-left row (0-based)
    ToCol   = 6,  // bottom-right column
    ToRow   = 22, // bottom-right row
    // ColOff / RowOff: fine-grained EMU offsets (default 0)
}
```

Multiple charts can be added to the same worksheet by calling `AddChart` more than once.

### ChartSeries

Use the fluent `.Series()` method to append series to a chart:

```
new ChartDefinition("Revenue by Product")
    .Type(ChartType.Column)
    .Legend("b")
    .Series(
        name:        "Revenue",           // label or sheet ref like "Sheet1!$B$1"
        valuesRef:   "Sheet1!$B$2:$B$6", // required: numeric data range
        categoryRef: "Sheet1!$A$2:$A$6") // optional: category labels
    .Series("Units", "Sheet1!$C$2:$C$6", "Sheet1!$A$2:$A$6"); // multi-series
```

### ODS Charts

`OdsFile` supports charts using the **same API** as `XlsxFile` — the same `ChartDefinition`,
`ChartType`, and `ObjectAnchor` types are used without modification.

```
// Works identically for both XlsxFile and OdsFile
sheet.AddChart(
    new ChartDefinition("Sales by Region")
        .Type(ChartType.Column)
        .Legend("b")
        .Series("Revenue",
                   valuesRef:   "Sales!$B$2:$B$6",
                   categoryRef: "Sales!$A$2:$A$6"),
    new ObjectAnchor { FromCol = 0, FromRow = 8, ToCol = 7, ToRow = 23 });
```

The library handles all ODS-specific details automatically:

| Detail | What happens internally |
|---|---|
| Cell references | Converted from XLSX style (`Sheet1!$A$2:$A$6`) to ODS style (`$Sheet1.$A$2:$Sheet1.$A$6`) |
| Chart sub-documents | Each chart is stored as an embedded ODS object (`Object 1/`, `Object 2/`, …) in the ZIP |
| Chart type mapping | `Column` → `chart:bar` (vertical), `Bar` → `chart:bar`, `Line` → `chart:line`, `Pie` → `chart:circle`, `Area` → `chart:area` |
| Legend positions | Mapped from XLSX codes to ODS names: `"b"` → `bottom`, `"t"` → `top`, `"l"` → `left`, `"r"` / `"tr"` → `right` |

All five chart types (`Column`, `Bar`, `Line`, `Pie`, `Area`) are supported in both XLSX and ODS output.

---

## Images

Embed raster images (PNG, JPEG, GIF, BMP, etc.) on any worksheet with `AddImage`.
Images are anchored using the same `ObjectAnchor` struct as charts.

### From a file path

```
// The content type is inferred from the file extension
sheet.AddImage(
    filePath: @"C:\assets\logo.png",
    anchor:   new ObjectAnchor { FromCol = 0, FromRow = 0, ToCol = 4, ToRow = 8 });
```

### From a byte array

```
byte[] imageBytes = File.ReadAllBytes("banner.jpg");

sheet.AddImage(
    imageBytes:  imageBytes,
    contentType: "image/jpeg",
    anchor:      new ObjectAnchor { FromCol = 0, FromRow = 0, ToCol = 6, ToRow = 12 });
```

### AddImage overloads

| Overload | Description |
|---|---|
| `AddImage(filePath, anchor?)` | Loads image from disk; content type inferred from extension |
| `AddImage(imageBytes, contentType, anchor?)` | Embeds raw bytes with the given MIME type |

Both overloads return `this` (`XlsxWorksheet`) for chaining.  
The `anchor` parameter is optional; when omitted the image is placed at the top-left corner (`FromCol = 0, FromRow = 0, ToCol = 5, ToRow = 10`).

### Supported MIME types / extensions

| MIME type | Extension |
|---|---|
| `image/png` | `.png` |
| `image/jpeg` | `.jpeg` / `.jpg` |
| `image/gif` | `.gif` |
| `image/bmp` | `.bmp` |
| `image/tiff` | `.tiff` / `.tif` |

Multiple images can be added to the same worksheet, and images and charts may coexist on the same sheet.

---

## CSV Export (`CsvFile`)

`CsvFile` provides a simple, flat API for producing RFC 4180–compliant CSV files.
Rows are added directly on the `CsvFile` instance — no worksheet abstraction is needed.

```
using FreeDataExportsv2;

var csv = new CsvFile
{
    Delimiter  = ",",    // default
    IncludeBom = true,   // UTF-8 BOM for Excel compatibility (default)
    LineEnding = "\r\n"  // CR+LF per RFC 4180 (default)
};

// Header row
csv.AddRow("Order #", "Product", "Qty", "Unit Price", "Sale Date", "Shipped");

// Data rows — pass any .NET values; formatting is automatic
foreach (var order in orders)
{
    csv.AddRow(
        order.Id,        // int   → "1001"
        order.Product,   // string as-is
        order.Qty,       // int   → "5"
        order.Price,     // decimal → "29.99"
        order.Date,      // DateTime → "1/5/2026"  (M/d/yyyy)
        order.Shipped);  // bool  → "TRUE" / "FALSE"
}

await csv.SaveAsync("orders.csv");
```

`AddRow` returns `this` so calls can be chained:

```
csv.AddRow("A", "B", "C")
   .AddRow(1, 2, 3)
   .AddRow(4, 5, 6);
```

### Auto-formatting rules

No `DataType` is required. Values are formatted automatically:

| .NET type | CSV output |
|---|---|
| `null` | *(empty)* |
| `bool` | `TRUE` / `FALSE` |
| `DateTime` | `M/d/yyyy` — e.g. `1/5/2026` |
| `DateTimeOffset` | date portion, `M/d/yyyy` |
| `decimal` / `double` / `float` | invariant-culture string |
| anything else | `ToString()` via invariant culture |

Values are automatically quoted (RFC 4180) when they contain the delimiter, a
double-quote character, or a newline. Internal double-quotes are escaped as `""`.

### CsvFile methods

| Method | Description |
|---|---|
| `AddRow(params object?[] values)` | Appends a row of auto-formatted values; returns `this`. |
| `GetBytes()` | Returns the CSV as a `byte[]` (with optional BOM). |
| `GetBytesAsync()` | Async version of `GetBytes()`. |
| `Save(path)` | Synchronous save to a file path. |
| `Save(stream)` | Synchronous write to a `Stream`. |
| `SaveAsync(path)` | Asynchronous save to a file path. |
| `SaveAsync(stream)` | Asynchronous write to a `Stream`. |

### CsvFile properties

| Property | Type | Default | Description |
|---|---|---|---|
| `Delimiter` | `string` | `","` | Field separator character(s). |
| `IncludeBom` | `bool` | `true` | Prepend a UTF-8 BOM (improves Excel auto-detection). |
| `LineEnding` | `string` | `"\r\n"` | Row separator — CR+LF per RFC 4180. |

---

## Error Handling

Every `AddCell` call is wrapped in a `try/catch`. When a conversion fails:

- The offending cell is left **empty with a red outline border**.
- The error (sheet name, cell reference, attempted value, exception message) is recorded.

### Display options

```
// Option A — Auto-create a red "Errors" worksheet on Save (only if errors exist)
workbook.AddErrorsWorksheet();

// Use a custom tab name
workbook.AddErrorsWorksheet("Data Issues");

// Option B — Retrieve errors as a formatted string at any time
string report = workbook.GetErrors();
if (!string.IsNullOrEmpty(report))
    Console.WriteLine(report);

// Both options can be used together
```

### Triggering an error deliberately

```
// Passing a non-date string where a date is expected logs the error gracefully
sheet.AddRow().AddCell("not-a-date", DataType.ShortDate);
```

---

## Saving / Getting Bytes

```
// Synchronous — file
workbook.Save("output.xlsx");

// Synchronous — stream
using var stream = File.Create("output.xlsx");
workbook.Save(stream);

// Asynchronous — file
await workbook.SaveAsync("output.xlsx");

// Asynchronous — stream
await workbook.SaveAsync(stream);

// In-memory byte array (e.g. for HTTP response)
byte[] bytes = workbook.GetBytes();

// Asynchronous byte array
byte[] bytes = await workbook.GetBytesAsync();
```

---

## DataType Reference

```
General        Number         Integer        Currency       Accounting
Percentage     WholePercent   Fraction       Scientific     Thousands
Thousands2     ShortDate      LongDate       DateTime       DateTime24
Time12h        Time24h        Boolean        Formula        Error
PhoneUS        Text           Zip
```

---

## CellOptions Reference

```
new CellOptions
{
    // Format
    DataType        = DataType.Currency,

    // Font
    FontName        = "Arial",       // null = Calibri (workbook default)
    FontSize        = 12.0,          // null = 11pt (workbook default)
    FontColor       = "FFFF0000",    // ARGB — null = default theme colour
    Bold            = true,
    Italic          = false,
    Underline       = false,
    Strikethrough   = false,

    // Fill
    BackgroundColor = "FFFFFF00",    // ARGB — null = no background

    // Alignment
    HorizontalAlign = "center",      // left | center | right | fill | justify
    VerticalAlign   = "bottom",      // top  | center | bottom | justify
    WrapText        = false,

    // Border — omit any side to leave it unstyled; color defaults to black
    BorderLeftStyle   = "thin",      // thin | medium | thick | dashed | dotted | double | hair | …
    BorderLeftColor   = "FF000000",  // ARGB — null = black
    BorderRightStyle  = "thin",
    BorderRightColor  = "FF000000",
    BorderTopStyle    = "thin",
    BorderTopColor    = "FF000000",
    BorderBottomStyle = "medium",
    BorderBottomColor = "FF1A3A5C",  // navy bottom accent
}
```

---

## XlsxTableStyles Reference

Use the `XlsxTableStyles` static class for type-safe style names, or pass any valid Excel
table style string directly.

```
XlsxTableStyles.Light1  … Light21
XlsxTableStyles.Medium1 … Medium28
XlsxTableStyles.Dark1   … Dark11
```

```
// Using a constant
.Style(XlsxTableStyles.Medium7)

// Using a raw string
.Style("TableStyleDark4")
```

---

## License

This project is licensed under the MIT License.  
See the [LICENSE](LICENSE) file for details.
