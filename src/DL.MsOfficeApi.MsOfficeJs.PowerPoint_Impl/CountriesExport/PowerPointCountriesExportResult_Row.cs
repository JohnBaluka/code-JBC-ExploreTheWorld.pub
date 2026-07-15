using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    // Result of insertCountries() in wwwroot/js/countries-export.js.
    // FlagError is set when the slides were added but the flag images could not be inserted.
    public record PowerPointCountriesExportResult_Row(
        [property: JsonPropertyName("inserted")] bool Inserted,
        [property: JsonPropertyName("count")] int Count,
        [property: JsonPropertyName("flagCount")] int FlagCount = 0,
        [property: JsonPropertyName("flagError")] string FlagError = "",
        [property: JsonPropertyName("error")] string Error = "");
}
