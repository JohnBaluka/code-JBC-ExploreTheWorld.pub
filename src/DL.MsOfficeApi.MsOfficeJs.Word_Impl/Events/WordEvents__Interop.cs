using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    public class WordEvents__Interop : JsModuleInterop__Base, WordEvents__Interop__Interface
    {
        public WordEvents__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl/js/events.js")
        {
        }

        public async Task StartWatchingAsync(object dotNetRef, string[] eventKeys)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("startWatching", dotNetRef, eventKeys);
        }

        public async Task StopWatchingAsync()
        {
            if (!IsInitialized) return;
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("stopWatching");
        }
    }
}
