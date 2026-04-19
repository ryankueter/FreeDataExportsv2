namespace FreeDataExportsv2.Internal;

internal sealed class XlsxAppProperties
{
    public string Application      { get; set; } = "Microsoft Office Excel";
    public string AppVersion       { get; set; } = "16.0300";
    public string Company          { get; set; } = string.Empty;
    public int    DocSecurity      { get; set; } = 0;
    public bool   ScaleCrop        { get; set; } = false;
    public bool   LinksUpToDate    { get; set; } = false;
    public bool   SharedDoc        { get; set; } = false;
    public bool   HyperlinksChanged { get; set; } = false;
}
