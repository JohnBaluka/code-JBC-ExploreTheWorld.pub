using JBC.ExploreTheWorld.CL;

namespace JBC.ExploreTheWorld.DL
{
    /// <summary>
    /// A resolved country flag image: the PNG bytes, the optional local cache file path
    /// (file-backed stores only), and whether it came from the cache or was downloaded.
    /// </summary>
    public class FlagImage_Row
    {
        public string Iso2 { get; set; } = string.Empty;
        public byte[] ImageBytes { get; set; } = System.Array.Empty<byte>();
        public string? FilePath { get; set; }
        public DataSource_Enum Source { get; set; }
    }
}
