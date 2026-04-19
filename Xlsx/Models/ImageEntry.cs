using FreeDataExportsv2;

namespace FreeDataExportsv2.Internal;

/// <summary>
/// Internal binding for an image embedded in a worksheet.
/// Created by <see cref="FreeDataExportsv2.XlsxWorksheet.AddImage"/> and
/// populated with global IDs by <see cref="FreeDataExportsv2.XlsxFile"/> before writing.
/// </summary>
internal sealed class XlsxImageEntry
{
    /// <summary>Raw image bytes (PNG, JPEG, GIF, BMP …).</summary>
    public byte[] ImageBytes { get; }

    /// <summary>MIME content type, e.g. <c>"image/png"</c>.</summary>
    public string ContentType { get; }

    /// <summary>Two-cell anchor that controls position and size on the sheet.</summary>
    public ObjectAnchor Anchor { get; }

    /// <summary>
    /// Global 1-based media ID assigned by <see cref="XlsxFile"/>.
    /// Determines the file name: <c>xl/media/image{MediaId}.{Extension}</c>.
    /// </summary>
    public int MediaId { get; set; }

    /// <summary>
    /// 1-based relationship ID used inside the drawing XML and its .rels file.
    /// Assigned by <see cref="XlsxFile"/> after chart rIds are allocated.
    /// </summary>
    public int DrawingRId { get; set; }

    /// <summary>File extension derived from <see cref="ContentType"/>.</summary>
    public string Extension => ContentType switch
    {
        "image/jpeg" or "image/jpg" => "jpeg",
        "image/gif"                 => "gif",
        "image/bmp"                 => "bmp",
        "image/tiff"                => "tiff",
        "image/webp"                => "webp",
        _                           => "png",
    };

    public XlsxImageEntry(byte[] imageBytes, string contentType, ObjectAnchor anchor)
    {
        ImageBytes  = imageBytes;
        ContentType = contentType;
        Anchor      = anchor;
    }
}
