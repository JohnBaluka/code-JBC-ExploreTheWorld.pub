using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    // Result of getDocumentAsJson() in wwwroot/js/save-as-json.js.
    public record WordSaveAsJsonResult_Row(
        [property: JsonPropertyName("json")] string Json,
        [property: JsonPropertyName("error")] string Error,
        [property: JsonPropertyName("fileName")] string FileName = "");
}
