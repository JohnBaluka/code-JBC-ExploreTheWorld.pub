using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Shape_Fields
    {
        public string? Name { get; set; }
        public int? Type { get; set; }
        public string? Left { get; set; }
        public string? Top { get; set; }
        public string? Width { get; set; }
        public string? Height { get; set; }
        public string? Rotation { get; set; }
        public string? AlternativeText { get; set; }
        public bool? HasTextFrame { get; set; }
        public string? TextFrameText { get; set; }
    }
}
