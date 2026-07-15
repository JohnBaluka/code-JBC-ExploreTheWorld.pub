using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    // Raw result of getPresentationData() in wwwroot/js/save-as-json.js. DataJson holds
    // the Office.js-shaped object graph (see MsPowerPointJs/PowerPointPresentationJs_Row).
    public record PowerPointSaveAsJsonData_Row(
        [property: JsonPropertyName("dataJson")] string DataJson,
        [property: JsonPropertyName("error")] string Error,
        [property: JsonPropertyName("fileName")] string FileName = "");
}
