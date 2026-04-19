using System.Xml.Linq;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Generates [Content_Types].xml — the OPC content-type registry at the root of the zip.
/// Every part in the package must have an entry here for Excel to recognise it.
/// </summary>
internal static class XlsxContentTypesPart
{
    private static readonly XNamespace Ns =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    public static byte[] Generate(int sheetCount,
                                   IReadOnlyList<int>?    tableIds        = null,
                                   IReadOnlyList<int>?    drawingIds      = null,
                                   IReadOnlyList<int>?    chartIds        = null,
                                   IReadOnlyList<string>? imageExtensions = null)
    {
        var elements = new List<object>
        {
            // Default catch-alls
            Default("rels", "application/vnd.openxmlformats-package.relationships+xml"),
            Default("xml",  "application/xml"),

            // XlsxWorkbook
            Override("/xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        };

        // Worksheets
        for (int i = 1; i <= sheetCount; i++)
            elements.Add(Override($"/xl/worksheets/sheet{i}.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));

        // Tables
        if (tableIds is not null)
            foreach (var id in tableIds)
                elements.Add(Override($"/xl/tables/table{id}.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"));

        // Drawings
        if (drawingIds is not null)
            foreach (var id in drawingIds)
                elements.Add(Override($"/xl/drawings/drawing{id}.xml",
                    "application/vnd.openxmlformats-officedocument.drawing+xml"));

        // Charts
        if (chartIds is not null)
            foreach (var id in chartIds)
                elements.Add(Override($"/xl/charts/chart{id}.xml",
                    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"));

        // Images — one Default entry per distinct file extension
        if (imageExtensions is not null)
            foreach (var ext in imageExtensions.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var mime = ext.ToLowerInvariant() switch
                {
                    "jpeg" or "jpg" => "image/jpeg",
                    "gif"           => "image/gif",
                    "bmp"           => "image/bmp",
                    "tiff"          => "image/tiff",
                    "webp"          => "image/webp",
                    _               => "image/png",
                };
                elements.Add(Default(ext, mime));
            }

        elements.Add(Override("/xl/theme/theme1.xml",
            "application/vnd.openxmlformats-officedocument.theme+xml"));
        elements.Add(Override("/xl/styles.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"));
        elements.Add(Override("/docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml"));
        elements.Add(Override("/docProps/app.xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml"));

        var root = new XElement(Ns + "Types", elements);
        return XlsxXmlHelper.ToXmlBytes(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }

    private static XElement Default(string extension, string contentType) =>
        new(Ns + "Default",
            new XAttribute("Extension",   extension),
            new XAttribute("ContentType", contentType));

    private static XElement Override(string partName, string contentType) =>
        new(Ns + "Override",
            new XAttribute("PartName",    partName),
            new XAttribute("ContentType", contentType));
}
