using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class TableCell_Fields
    {
        public long? RowIndex { get; set; }
        public long? ColumnIndex { get; set; }
        public string? Text { get; set; }
    }
}
