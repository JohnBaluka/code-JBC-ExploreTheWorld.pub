using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class InlineShape : InlineShape_Fields
    {
        // Entities
        public ImageBlob? Image { get; set; }
    }
}
