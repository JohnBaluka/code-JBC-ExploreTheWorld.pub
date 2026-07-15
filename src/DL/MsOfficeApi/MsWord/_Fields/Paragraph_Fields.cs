using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Paragraph_Fields
    {
        public string? Text { get; set; }
        public string? Style { get; set; }
        public int? Alignment { get; set; }
        public string? LeftIndent { get; set; }
        public string? RightIndent { get; set; }
        public string? FirstLineIndent { get; set; }
        public string? SpaceBefore { get; set; }
        public string? SpaceAfter { get; set; }
        public string? LineSpacing { get; set; }
        public bool? KeepWithNext { get; set; }
        public bool? PageBreakBefore { get; set; }
    }
}
