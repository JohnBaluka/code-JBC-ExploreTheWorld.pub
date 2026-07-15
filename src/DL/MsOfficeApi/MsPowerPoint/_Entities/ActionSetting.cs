using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class ActionSetting : ActionSetting_Fields
    {
        // Entities
        public Hyperlink? Hyperlink { get; set; }
        public SoundEffect? SoundEffect { get; set; }
    }
}
