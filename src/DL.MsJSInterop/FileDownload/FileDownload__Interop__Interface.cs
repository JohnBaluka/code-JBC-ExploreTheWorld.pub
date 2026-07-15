namespace JBC.ExploreTheWorld.DL.MsJSInterop.FileDownload
{
    // Generic browser file-download interop (wraps wwwroot/js/download-file.js).
    public interface FileDownload__Interop__Interface
    {
        Task DownloadFileFromBytesAsync(string fileName, string contentType, byte[] bytes);
        Task DownloadTextAsync(string fileName, string contentType, string text);
    }
}
