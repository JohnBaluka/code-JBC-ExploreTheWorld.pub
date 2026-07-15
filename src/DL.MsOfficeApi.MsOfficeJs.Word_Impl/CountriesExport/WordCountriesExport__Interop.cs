using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    public class WordCountriesExport__Interop : JsModuleInterop__Base, WordCountriesExport__Interop__Interface
    {
        public WordCountriesExport__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl/js/countries-export.js")
        {
        }

        // payloadJson: { "countries": [ { "country", "iso2", "iso3", "flagUrl" }, ... ] }
        public async Task<WordCountriesExportResult_Row> InsertCountriesAsync(string payloadJson)
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<WordCountriesExportResult_Row>("insertCountries", payloadJson);
        }
    }
}
