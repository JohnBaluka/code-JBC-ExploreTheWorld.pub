using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Section : Section_Fields
    {
        // Entities
        public PageSetup? PageSetup { get; set; }
    }
}
