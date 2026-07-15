using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class ColorFormat_Fields
    {
        public string? Brightness { get; set; }
        public long? Creator { get; set; }
        public int? ObjectThemeColor { get; set; }
        public int? RGB { get; set; }
        public int? SchemeColor { get; set; }
        public string? TintAndShade { get; set; }
        public int? Type { get; set; }
    }
}
