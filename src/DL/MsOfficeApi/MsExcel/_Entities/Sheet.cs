using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class Sheet : Sheet_Fields
    {
        // Entities
        public UsedRange? UsedRange { get; set; }

        // Lists
        public List<Row>? Rows { get; set; }
    }
}
