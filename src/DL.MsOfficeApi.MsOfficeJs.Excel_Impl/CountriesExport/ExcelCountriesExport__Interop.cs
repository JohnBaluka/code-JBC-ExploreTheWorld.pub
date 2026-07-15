using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    public class ExcelCountriesExport__Interop : JsModuleInterop__Base, ExcelCountriesExport__Interop__Interface
    {
        public ExcelCountriesExport__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl/js/countries-export.js")
        {
        }

        // payloadJson: { "countries": [ { "country", "iso2", "iso3", "flagUrl" }, ... ] }
        public async Task<ExcelCountriesExportResult_Row> InsertCountriesAsync(string payloadJson)
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<ExcelCountriesExportResult_Row>("insertCountries", payloadJson);
        }
    }
}
