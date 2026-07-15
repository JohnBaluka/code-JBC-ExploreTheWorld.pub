using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    // Result of getPresentationAsJson() in wwwroot/js/save-as-json.js.
    public record PowerPointSaveAsJsonResult_Row(
        [property: JsonPropertyName("json")] string Json,
        [property: JsonPropertyName("error")] string Error,
        [property: JsonPropertyName("fileName")] string FileName = "");
}
