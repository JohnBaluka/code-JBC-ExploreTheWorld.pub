namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    // Wraps wwwroot/js/save-as-json.js (serialize active Excel workbook to JSON + download).
    public interface ExcelSaveAsJson__Interop__Interface
    {
        Task<ExcelSaveAsJsonResult_Row> GetWorkbookAsJsonAsync();
        Task DownloadJsonAsync(string json, string fileName);
    }
}
