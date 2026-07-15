using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL
{
    /// <summary>Downloads a flag image (PNG bytes) from an absolute URL.</summary>
    public interface FlagImageDownload__Repo__Interface
    {
        /// <summary>Downloads the image at <paramref name="url"/> and returns its bytes.</summary>
        Task<byte[]> DownloadImageAsync(string url);
    }
}
