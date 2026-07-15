using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    // Raw result of getDocumentData() in wwwroot/js/save-as-json.js. DataJson holds the
    // Office.js-shaped object graph (see MsWordJs/WordDocumentJs_Row).
    public record WordSaveAsJsonData_Row(
        [property: JsonPropertyName("dataJson")] string DataJson,
        [property: JsonPropertyName("error")] string Error,
        [property: JsonPropertyName("fileName")] string FileName = "");
}
