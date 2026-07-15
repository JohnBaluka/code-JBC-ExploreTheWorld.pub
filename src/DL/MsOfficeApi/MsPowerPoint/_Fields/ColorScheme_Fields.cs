using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class ColorScheme_Fields
    {
        public string? Accent1 { get; set; }
        public string? Accent2 { get; set; }
        public string? Accent3 { get; set; }
        public string? Background { get; set; }
        public string? Fill { get; set; }
        public string? Foreground { get; set; }
        public string? NotSchemeColor { get; set; }
        public string? SchemeColorMixed { get; set; }
        public string? Shadow { get; set; }
        public string? Title { get; set; }

    }
}
