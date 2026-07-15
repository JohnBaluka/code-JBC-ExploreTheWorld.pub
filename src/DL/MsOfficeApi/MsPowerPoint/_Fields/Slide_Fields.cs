using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Slide_Fields
    {
        public string? Name { get; set; }
        public string? CustomLayout_Name { get; set; }
        public string? Design_Name { get; set; }
        public int? SlideID { get; set; }
        public int? SlideIndex { get; set; }
        public int? SlideNumber { get; set; }
        public int? sectionIndex { get; set; }
        public int? BackgroundStyle { get; set; }
        public int? DisplayMasterShapes { get; set; }
        public int? FollowMasterBackground { get; set; }
        public int? HasNotesPage { get; set; }
        public int? Layout { get; set; }
        public long? PrintSteps { get; set; }
    }
}
