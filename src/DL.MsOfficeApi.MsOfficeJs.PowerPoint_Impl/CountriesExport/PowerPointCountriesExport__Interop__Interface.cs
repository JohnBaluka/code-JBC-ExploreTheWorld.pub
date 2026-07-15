namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    // Wraps wwwroot/js/countries-export.js — appends one slide per loaded CountriesNow row
    // (Country / ISO2 / ISO3 text plus the flag image) to the active presentation via Office.js.
    public interface PowerPointCountriesExport__Interop__Interface
    {
        Task<PowerPointCountriesExportResult_Row> InsertCountriesAsync(string payloadJson);
    }
}
