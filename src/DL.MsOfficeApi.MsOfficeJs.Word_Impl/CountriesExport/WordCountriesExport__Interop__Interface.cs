namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    // Wraps wwwroot/js/countries-export.js — appends the loaded CountriesNow rows to the
    // active Word document as a table (Flag, Country, ISO2, ISO3) via the Office.js API.
    public interface WordCountriesExport__Interop__Interface
    {
        Task<WordCountriesExportResult_Row> InsertCountriesAsync(string payloadJson);
    }
}
