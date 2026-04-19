namespace FreeDataExportsv2.Internal;

internal sealed class XlsxCoreProperties
{
    public string   Creator        { get; set; } = string.Empty;
    public string   LastModifiedBy { get; set; } = string.Empty;
    public DateTime Created        { get; set; } = DateTime.UtcNow;
    public DateTime Modified       { get; set; } = DateTime.UtcNow;
    public string?  Title          { get; set; }
    public string?  Subject        { get; set; }
    public string?  Description    { get; set; }
    public string?  Keywords       { get; set; }
}
