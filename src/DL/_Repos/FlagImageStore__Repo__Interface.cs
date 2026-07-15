using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL
{
    /// <summary>
    /// Local cache store for country flag images keyed by ISO2 code. Implementations:
    /// file system (desktop/server hosts) and browser IndexedDB (WASM/web add-in hosts).
    /// </summary>
    public interface FlagImageStore__Repo__Interface
    {
        /// <summary>Returns the cached image bytes for the ISO2 code, or <c>null</c> when not cached.</summary>
        Task<byte[]?> GetImageAsync(string iso2);

        /// <summary>Saves the image bytes for the ISO2 code, overwriting any existing entry.</summary>
        Task SaveImageAsync(string iso2, byte[] imageBytes);

        /// <summary>
        /// Returns the local file path of the cached image when the store is file-backed
        /// (used by COM/VBA writers that require a path), or <c>null</c> otherwise.
        /// </summary>
        Task<string?> GetImageFilePathAsync(string iso2);

        /// <summary>Removes all cached flag images.</summary>
        Task ClearAsync();
    }
}
