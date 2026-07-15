using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class SlideRange : SlideRange_Fields
    {
        // Lists
        public List<Shape>? Shapes { get; set; }
    }
}
