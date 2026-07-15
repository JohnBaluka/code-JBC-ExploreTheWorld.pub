using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    // Result of insertCountries() in wwwroot/js/countries-export.js.
    public record WordCountriesExportResult_Row(
        [property: JsonPropertyName("inserted")] bool Inserted,
        [property: JsonPropertyName("count")] int Count,
        [property: JsonPropertyName("flagCount")] int FlagCount = 0,
        [property: JsonPropertyName("error")] string Error = "");
}
