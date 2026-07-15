using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    public class PowerPointEvents__Interop : JsModuleInterop__Base, PowerPointEvents__Interop__Interface
    {
        public PowerPointEvents__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl/js/events.js")
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
