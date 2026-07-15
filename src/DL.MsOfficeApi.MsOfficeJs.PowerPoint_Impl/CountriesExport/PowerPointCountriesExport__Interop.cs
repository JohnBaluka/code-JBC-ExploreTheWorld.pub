using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    public class PowerPointCountriesExport__Interop : JsModuleInterop__Base, PowerPointCountriesExport__Interop__Interface
    {
        public PowerPointCountriesExport__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl/js/countries-export.js")
        {
        }

        // payloadJson: { "countries": [ { "country", "iso2", "iso3", "flagUrl" }, ... ] }
        public async Task<PowerPointCountriesExportResult_Row> InsertCountriesAsync(string payloadJson)
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<PowerPointCountriesExportResult_Row>("insertCountries", payloadJson);
        }
    }
}
