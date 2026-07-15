using System.Text.Json;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    public class ExcelSaveAsJson__Interop : JsModuleInterop__Base, ExcelSaveAsJson__Interop__Interface
    {
        public ExcelSaveAsJson__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl/js/save-as-json.js")
        {
        }

        // Collects the Office.js object graph, maps it to the canonical
        // JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel entities and serializes it with
        // MsOfficeJsonSerializer so the output matches the other writers.
        public async Task<ExcelSaveAsJsonResult_Row> GetWorkbookAsJsonAsync()
        {
            var module = await GetModuleAsync();
            var data = await module.InvokeAsync<ExcelSaveAsJsonData_Row>("getWorkbookData");

            if (!string.IsNullOrEmpty(data.Error))
            {
                return new ExcelSaveAsJsonResult_Row(string.Empty, data.Error, data.FileName);
            }

            var jsRow = JsonSerializer.Deserialize<ExcelWorkbookJs_Row>(data.DataJson)
                ?? new ExcelWorkbookJs_Row();

            var workbook = MsExcelJsMapper.ToWorkbook(jsRow, data.FileName);

            string json = MsOfficeJsonSerializer.Serialize(workbook) + "\r\n";

            return new ExcelSaveAsJsonResult_Row(json, string.Empty, data.FileName);
        }

        public async Task DownloadJsonAsync(string json, string fileName)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("downloadJson", json, fileName);
        }
    }
}
