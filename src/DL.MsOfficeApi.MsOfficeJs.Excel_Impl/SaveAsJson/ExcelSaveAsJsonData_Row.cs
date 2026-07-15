using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    // Raw result of getWorkbookData() in wwwroot/js/save-as-json.js. DataJson holds the
    // Office.js-shaped object graph (see MsExcelJs/ExcelWorkbookJs_Row).
    public record ExcelSaveAsJsonData_Row(
        [property: JsonPropertyName("dataJson")] string DataJson,
        [property: JsonPropertyName("error")] string Error,
        [property: JsonPropertyName("fileName")] string FileName = "");
}
