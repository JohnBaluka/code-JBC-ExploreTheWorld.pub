using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class Cell_Fields
    {
        public long? RowIndex { get; set; }
        public long? ColumnIndex { get; set; }
        public string? Address { get; set; }
        public string? Value { get; set; }
        public string? Formula { get; set; }
        public string? NumberFormat { get; set; }
        public string? Text { get; set; }
        public int? HorizontalAlignment { get; set; }
        public int? VerticalAlignment { get; set; }
        public bool? WrapText { get; set; }
        public bool? MergeCells { get; set; }
        public bool? HasComment { get; set; }
        public string? Comment { get; set; }
    }
}
