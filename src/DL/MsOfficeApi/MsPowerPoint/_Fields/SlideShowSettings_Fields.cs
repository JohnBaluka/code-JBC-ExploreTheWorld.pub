using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class SlideShowSettings_Fields
    {
        public int? AdvanceMode { get; set; }
        public long? EndingSlide { get; set; }
        public int? LoopUntilStopped { get; set; }
        public int? RangeType { get; set; }
        public int? ShowMediaControls { get; set; }
        public int? ShowPresenterView { get; set; }
        public int? ShowScrollbar { get; set; }
        public int? ShowType { get; set; }
        public int? ShowWithAnimation { get; set; }
        public int? ShowWithNarration { get; set; }
        public long? StartingSlide { get; set; }
    }
}
