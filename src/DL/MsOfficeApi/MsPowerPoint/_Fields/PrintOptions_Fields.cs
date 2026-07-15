using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class PrintOptions_Fields
    {
        public string? ActivePrinter { get; set; }
        public int? Collate { get; set; }
        public int? FitToPage { get; set; }
        public int? FrameSlides { get; set; }
        public int? HandoutOrder { get; set; }
        public int? HighQuality { get; set; }
        public long? NumberOfCopies { get; set; }
        public int? OutputType { get; set; }
        public int? PrintColorType { get; set; }
        public int? PrintComments { get; set; }
        public int? PrintFontsAsGraphics { get; set; }
        public int? PrintHiddenSlides { get; set; }
        public int? PrintInBackground { get; set; }
        public int? RangeType { get; set; }
        public string? SlideShowName { get; set; }
    }
}
