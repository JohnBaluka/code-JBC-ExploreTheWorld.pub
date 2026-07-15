using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsJSInterop.FlagImageCache
{
    public class FlagImageCache__Interop : JsModuleInterop__Base, FlagImageCache__Interop__Interface
    {
        public FlagImageCache__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsJSInterop/js/flag-image-cache.js")
        {
        }

        public async Task<byte[]?> GetImageAsync(string iso2)
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<byte[]?>("getImage", iso2);
        }

        public async Task SaveImageAsync(string iso2, byte[] imageBytes)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("saveImage", iso2, imageBytes);
        }

        public async Task ClearImagesAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("clearImages");
        }
    }
}
