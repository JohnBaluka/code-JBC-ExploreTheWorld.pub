using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Design_Fields
    {
        public string? Name { get; set; }
        public long? Index { get; set; }
        public int? HasTitleMaster { get; set; }
        public int? Preserved { get; set; }
    }
}
