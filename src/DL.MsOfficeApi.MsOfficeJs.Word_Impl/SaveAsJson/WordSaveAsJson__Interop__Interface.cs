namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    // Wraps wwwroot/js/save-as-json.js (serialize active Word document to JSON + download).
    public interface WordSaveAsJson__Interop__Interface
    {
        Task<WordSaveAsJsonResult_Row> GetDocumentAsJsonAsync();
        Task DownloadJsonAsync(string json, string fileName);
    }
}
