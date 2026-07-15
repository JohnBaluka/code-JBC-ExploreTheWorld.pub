namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    // Wraps wwwroot/js/save-as-json.js (serialize active presentation to JSON + download).
    public interface PowerPointSaveAsJson__Interop__Interface
    {
        Task<PowerPointSaveAsJsonResult_Row> GetPresentationAsJsonAsync();
        Task DownloadJsonAsync(string json, string fileName);
    }
}
