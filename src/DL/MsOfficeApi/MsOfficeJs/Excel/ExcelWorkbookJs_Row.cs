using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel
{
    // JS-shaped rows matching the object graph collected by getWorkbookData() in
    // wwwroot/js/save-as-json.js (the Office.js Excel object model). These are mapped
    // to the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel entities by MsExcelJsMapper.
    public class ExcelWorkbookJs_Row
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("properties")]
        public ExcelDocumentPropertiesJs_Row? Properties { get; set; }

        [JsonPropertyName("sheets")]
        public List<ExcelSheetJs_Row> Sheets { get; set; } = new();
    }
}
