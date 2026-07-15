using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Style_Fields
    {
        public string? NameLocal { get; set; }
        public int? Type { get; set; }
        public bool? InUse { get; set; }
        public bool? BuiltIn { get; set; }
        public bool? AutomaticallyUpdate { get; set; }
    }
}
