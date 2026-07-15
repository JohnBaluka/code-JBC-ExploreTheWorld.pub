using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class PageSetup_Fields
    {
        public string? PageWidth { get; set; }
        public string? PageHeight { get; set; }
        public int? Orientation { get; set; }
        public string? TopMargin { get; set; }
        public string? BottomMargin { get; set; }
        public string? LeftMargin { get; set; }
        public string? RightMargin { get; set; }
        public string? HeaderDistance { get; set; }
        public string? FooterDistance { get; set; }
        public int? PaperSize { get; set; }
        public int? SectionStart { get; set; }
        public bool? DifferentFirstPageHeaderFooter { get; set; }
    }
}
