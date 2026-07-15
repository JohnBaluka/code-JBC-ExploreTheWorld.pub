using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    // Result of insertCountries() in wwwroot/js/countries-export.js.
    public record ExcelCountriesExportResult_Row(
        [property: JsonPropertyName("inserted")] bool Inserted,
        [property: JsonPropertyName("count")] int Count,
        [property: JsonPropertyName("flagCount")] int FlagCount = 0,
        [property: JsonPropertyName("error")] string Error = "");
}
