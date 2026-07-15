namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    // Wraps wwwroot/js/countries-export.js — writes the loaded CountriesNow rows to the
    // active worksheet (Flag, Country, ISO2, ISO3 starting at A1) via the Office.js API.
    public interface ExcelCountriesExport__Interop__Interface
    {
        Task<ExcelCountriesExportResult_Row> InsertCountriesAsync(string payloadJson);
    }
}
