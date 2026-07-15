namespace JBC.ExploreTheWorld.DL.MsJSInterop.FlagImageCache
{
    /// <summary>JS interop for the IndexedDB flag image cache (flag-image-cache.js).</summary>
    public interface FlagImageCache__Interop__Interface
    {
        Task<byte[]?> GetImageAsync(string iso2);
        Task SaveImageAsync(string iso2, byte[] imageBytes);
        Task ClearImagesAsync();
    }
}
