using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.DL.MsJSInterop.FlagImageCache
{
    /// <summary>
    /// Browser implementation of <see cref="FlagImageStore__Repo__Interface"/> backed by
    /// the IndexedDB flag-image cache interop. There is no local file path in the browser,
    /// so <see cref="GetImageFilePathAsync"/> always returns <c>null</c>.
    /// </summary>
    public class FlagImageStore_Browser__Repo : FlagImageStore__Repo__Interface
    {
        private readonly FlagImageCache__Interop__Interface _interop;

        public FlagImageStore_Browser__Repo(FlagImageCache__Interop__Interface interop)
        {
            _interop = interop;
        }

        public Task<byte[]?> GetImageAsync(string iso2) => _interop.GetImageAsync(iso2);

        public Task SaveImageAsync(string iso2, byte[] imageBytes) => _interop.SaveImageAsync(iso2, imageBytes);

        public Task<string?> GetImageFilePathAsync(string iso2) => Task.FromResult<string?>(null);

        public Task ClearAsync() => _interop.ClearImagesAsync();
    }
}
