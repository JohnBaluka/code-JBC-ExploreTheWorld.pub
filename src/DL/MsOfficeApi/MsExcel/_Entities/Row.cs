using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class Row : Row_Fields
    {
        // Lists
        public List<Cell>? Cells { get; set; }
    }
}
