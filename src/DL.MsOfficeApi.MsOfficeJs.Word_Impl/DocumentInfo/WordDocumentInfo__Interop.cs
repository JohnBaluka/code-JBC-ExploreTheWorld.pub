using JBC.ExploreTheWorld.DL.MsJSInterop;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    public class WordDocumentInfo__Interop : JsModuleInterop__Base, WordDocumentInfo__Interop__Interface
    {
        public WordDocumentInfo__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl/js/document-info.js")
        {
        }

        public async Task<WordDocumentInfo_Row> GetDocumentInfoAsync()
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<WordDocumentInfo_Row>("getDocumentInfo");
        }
    }
}
