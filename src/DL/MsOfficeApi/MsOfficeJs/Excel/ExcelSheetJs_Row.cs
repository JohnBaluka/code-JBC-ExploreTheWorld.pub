using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel
{
    public class ExcelSheetJs_Row
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        // Office.js worksheet position (zero-based).
        [JsonPropertyName("position")]
        public int Position { get; set; }

        // Office.js Excel.SheetVisibility value ("Visible", "Hidden", "VeryHidden").
        [JsonPropertyName("visibility")]
        public string? Visibility { get; set; }

        [JsonPropertyName("address")]
        public string? Address { get; set; }

        [JsonPropertyName("rowCount")]
        public long RowCount { get; set; }

        [JsonPropertyName("columnCount")]
        public long ColumnCount { get; set; }

        [JsonPropertyName("values")]
        public List<List<string?>>? Values { get; set; }

        [JsonPropertyName("formulas")]
        public List<List<string?>>? Formulas { get; set; }
    }
}
