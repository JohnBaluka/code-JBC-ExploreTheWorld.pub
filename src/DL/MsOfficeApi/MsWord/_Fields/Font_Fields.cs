using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Font_Fields
    {
        public string? Name { get; set; }
        public string? Size { get; set; }
        public int? Bold { get; set; }
        public int? Italic { get; set; }
        public int? Underline { get; set; }
        public long? Color { get; set; }
        public int? ColorIndex { get; set; }
        public int? StrikeThrough { get; set; }
        public int? Superscript { get; set; }
        public int? Subscript { get; set; }
        public string? Scaling { get; set; }
        public string? Spacing { get; set; }
        public string? Position { get; set; }
        public int? Emboss { get; set; }
        public int? Shadow { get; set; }
    }
}
