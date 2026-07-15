using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Paragraph : Paragraph_Fields
    {
        // Lists
        public List<Run>? Runs { get; set; }
    }
}
