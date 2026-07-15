using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsJSInterop.FileDownload
{
    public class FileDownload__Interop : JsModuleInterop__Base, FileDownload__Interop__Interface
    {
        public FileDownload__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsJSInterop/js/download-file.js")
        {
        }

        public async Task DownloadFileFromBytesAsync(string fileName, string contentType, byte[] bytes)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("downloadFileFromBytes", fileName, contentType, bytes);
        }

        public async Task DownloadTextAsync(string fileName, string contentType, string text)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("downloadText", fileName, contentType, text);
        }
    }
}
