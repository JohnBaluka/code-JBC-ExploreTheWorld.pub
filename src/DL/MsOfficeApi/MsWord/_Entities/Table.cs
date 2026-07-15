using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Table : Table_Fields
    {
        // Lists
        public List<TableRow>? Rows { get; set; }
    }
}
