using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class PlaySettings_Fields
    {
        public string? ActionVerb { get; set; }
        public int? HideWhileNotPlaying { get; set; }
        public int? LoopUntilStopped { get; set; }
        public int? PauseAnimation { get; set; }
        public int? PlayOnEntry { get; set; }
        public int? RewindMovie { get; set; }
        public long? StopAfterSlides { get; set; }
    }
}
