using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class TableRow : TableRow_Fields
    {
        // Lists
        public List<TableCell>? Cells { get; set; }
    }
}
