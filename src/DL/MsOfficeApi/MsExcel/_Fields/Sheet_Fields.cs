using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class Sheet_Fields
    {
        public string? Name { get; set; }
        public long? Index { get; set; }
        public string? CodeName { get; set; }
        public int? Visible { get; set; }
        public string? StandardWidth { get; set; }
        public string? Type { get; set; }
    }
}
