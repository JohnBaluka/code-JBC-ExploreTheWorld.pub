using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Bookmark_Fields
    {
        public string? Name { get; set; }
        public string? RangeText { get; set; }
        public bool? Empty { get; set; }
    }
}
