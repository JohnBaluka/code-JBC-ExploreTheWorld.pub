using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    // Result of getWorkbookInfo() in wwwroot/js/workbook-info.js.
    public class ExcelWorkbookInfo_Row
    {
        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; } = string.Empty;

        [JsonPropertyName("activeSheetName")]
        public string ActiveSheetName { get; set; } = string.Empty;

        [JsonPropertyName("sheetCount")]
        public int SheetCount { get; set; }

        [JsonPropertyName("usedRangeAddress")]
        public string UsedRangeAddress { get; set; } = string.Empty;

        [JsonPropertyName("errorMessage")]
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
