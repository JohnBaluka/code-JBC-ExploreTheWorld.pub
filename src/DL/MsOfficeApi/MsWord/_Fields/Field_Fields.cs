using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Field_Fields
    {
        public long? Index { get; set; }
        public int? Type { get; set; }
        public bool? Locked { get; set; }
        public string? Code { get; set; }
        public string? Result { get; set; }
    }
}
