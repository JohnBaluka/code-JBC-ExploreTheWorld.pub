using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class ContentControl_Fields
    {
        public string? Title { get; set; }
        public string? Tag { get; set; }
        public int? Type { get; set; }
        public string? RangeText { get; set; }
    }
}
