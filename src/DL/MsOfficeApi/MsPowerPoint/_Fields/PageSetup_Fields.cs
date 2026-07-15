using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class PageSetup_Fields
    {
        public long? FirstSlideNumber { get; set; }
        public int? NotesOrientation { get; set; }
        public string? SlideHeight { get; set; }
        public int? SlideOrientation { get; set; }
        public int? SlideSize { get; set; }
        public string? SlideWidth { get; set; }
    }
}
