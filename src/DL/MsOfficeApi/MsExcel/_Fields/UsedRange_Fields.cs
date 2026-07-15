using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class UsedRange_Fields
    {
        public string? Address { get; set; }
        public long? RowCount { get; set; }
        public long? ColumnCount { get; set; }
    }
}
