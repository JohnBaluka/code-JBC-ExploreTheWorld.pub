using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class Font_Fields
    {
        public string? Name { get; set; }
        public string? Size { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public int? Underline { get; set; }
        public long? Color { get; set; }
        public int? ColorIndex { get; set; }
        public bool? StrikeThrough { get; set; }
    }
}
