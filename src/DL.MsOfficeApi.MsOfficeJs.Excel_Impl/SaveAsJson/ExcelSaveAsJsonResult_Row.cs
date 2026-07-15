using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    // Result of getWorkbookAsJson() in wwwroot/js/save-as-json.js.
    public record ExcelSaveAsJsonResult_Row(
        [property: JsonPropertyName("json")] string Json,
        [property: JsonPropertyName("error")] string Error,
        [property: JsonPropertyName("fileName")] string FileName = "");
}
