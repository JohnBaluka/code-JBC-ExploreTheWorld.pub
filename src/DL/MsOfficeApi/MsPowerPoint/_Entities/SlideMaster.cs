using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class SlideMaster: SlideMaster_Fields
    {
        // Entities
        public SlideShowTransition? SlideShowTransition { get; set; }

        // Lists
        public List<SlideMasterLayout>? SlideMasterLayouts { get; set; }
        public List<Shape>? Shapes { get; set; }

    }
}
