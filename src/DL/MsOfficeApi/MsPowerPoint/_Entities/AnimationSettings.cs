using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class AnimationSettings : AnimationSettings_Fields
    {
        // Entities
        public ColorFormat? DimColor { get; set; }
        public PlaySettings? PlaySettings { get; set; }
        public SoundEffect? SoundEffect { get; set; }
    }
}
