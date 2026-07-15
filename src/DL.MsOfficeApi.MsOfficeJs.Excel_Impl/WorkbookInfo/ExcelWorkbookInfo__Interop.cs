using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    public class ExcelWorkbookInfo__Interop : JsModuleInterop__Base, ExcelWorkbookInfo__Interop__Interface
    {
        public ExcelWorkbookInfo__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl/js/workbook-info.js")
        {
        }

        public async Task<ExcelWorkbookInfo_Row> GetWorkbookInfoAsync()
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<ExcelWorkbookInfo_Row>("getWorkbookInfo");
        }
    }
}
