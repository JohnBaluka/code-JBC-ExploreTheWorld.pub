using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class SlideShowTransition_Fields
    {
        public int? AdvanceOnClick { get; set; }
        public int? AdvanceOnTime { get; set; }
        public string? AdvanceTime { get; set; }
        public string? Duration { get; set; }
        public int? EntryEffect { get; set; }
        public int? Hidden { get; set; }
        public int? LoopSoundUntilNext { get; set; }
        public int? Speed { get; set; }
    }
}
