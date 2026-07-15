using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsJSInterop.Layout
{
    public class Layout__Interop : JsModuleInterop__Base, Layout__Interop__Interface
    {
        public Layout__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsJSInterop/js/layout.js")
        {
        }

        public async Task<int> GetWindowWidthAsync()
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<int>("getWindowWidth");
        }

        public async Task<IJSObjectReference> WatchWindowWidthAsync(object dotNetRef)
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<IJSObjectReference>("watchWindowWidth", dotNetRef);
        }
    }
}
