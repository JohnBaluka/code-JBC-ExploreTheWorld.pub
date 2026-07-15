using System.Text.Json;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    public class PowerPointSaveAsJson__Interop : JsModuleInterop__Base, PowerPointSaveAsJson__Interop__Interface
    {
        public PowerPointSaveAsJson__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl/js/save-as-json.js")
        {
        }

        // Collects the Office.js object graph, maps it to the canonical
        // JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint entities and serializes it with
        // MsOfficeJsonSerializer so the output matches the other writers.
        public async Task<PowerPointSaveAsJsonResult_Row> GetPresentationAsJsonAsync()
        {
            var module = await GetModuleAsync();
            var data = await module.InvokeAsync<PowerPointSaveAsJsonData_Row>("getPresentationData");

            if (!string.IsNullOrEmpty(data.Error))
            {
                return new PowerPointSaveAsJsonResult_Row(string.Empty, data.Error, data.FileName);
            }

            var jsRow = JsonSerializer.Deserialize<PowerPointPresentationJs_Row>(data.DataJson)
                ?? new PowerPointPresentationJs_Row();

            var presentation = MsPowerPointJsMapper.ToPresentation(jsRow, data.FileName);

            string json = MsOfficeJsonSerializer.Serialize(presentation) + "\r\n";

            return new PowerPointSaveAsJsonResult_Row(json, string.Empty, data.FileName);
        }

        public async Task DownloadJsonAsync(string json, string fileName)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("downloadJson", json, fileName);
        }
    }
}
