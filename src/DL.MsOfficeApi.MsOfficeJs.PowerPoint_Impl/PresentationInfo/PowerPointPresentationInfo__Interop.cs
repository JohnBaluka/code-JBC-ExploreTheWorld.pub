using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    public class PowerPointPresentationInfo__Interop : JsModuleInterop__Base, PowerPointPresentationInfo__Interop__Interface
    {
        public PowerPointPresentationInfo__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl/js/presentation-info.js")
        {
        }

        public async Task<PowerPointPresentationInfo_Row> GetPresentationInfoAsync()
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<PowerPointPresentationInfo_Row>("getPresentationInfo");
        }
    }
}
