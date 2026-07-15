using System.Text.Json;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    public class WordSaveAsJson__Interop : JsModuleInterop__Base, WordSaveAsJson__Interop__Interface
    {
        public WordSaveAsJson__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl/js/save-as-json.js")
        {
        }

        // Collects the Office.js object graph, maps it to the canonical
        // JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord entities and serializes it with
        // MsOfficeJsonSerializer so the output matches the other writers.
        public async Task<WordSaveAsJsonResult_Row> GetDocumentAsJsonAsync()
        {
            var module = await GetModuleAsync();
            var data = await module.InvokeAsync<WordSaveAsJsonData_Row>("getDocumentData");

            if (!string.IsNullOrEmpty(data.Error))
            {
                return new WordSaveAsJsonResult_Row(string.Empty, data.Error, data.FileName);
            }

            var jsRow = JsonSerializer.Deserialize<WordDocumentJs_Row>(data.DataJson)
                ?? new WordDocumentJs_Row();

            var document = MsWordJsMapper.ToDocument(jsRow, data.FileName);

            string json = MsOfficeJsonSerializer.Serialize(document) + "\r\n";

            return new WordSaveAsJsonResult_Row(json, string.Empty, data.FileName);
        }

        public async Task DownloadJsonAsync(string json, string fileName)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("downloadJson", json, fileName);
        }
    }
}
