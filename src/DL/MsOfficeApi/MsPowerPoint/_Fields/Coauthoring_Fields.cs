using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Coauthoring_Fields
    {
        public bool? FavorServerEditsDuringMerge { get; set; }
        public bool? MergeMode { get; set; }
        public bool? PendingUpdates { get; set; }
    }
}
