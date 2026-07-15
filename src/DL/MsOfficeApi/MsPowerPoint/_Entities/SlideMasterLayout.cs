using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class SlideMasterLayout : SlideMasterLayout_Fields
    {
        // Entities
        public SlideShowTransition? SlideShowTransition { get; set; }

        // Lists
        public List<Shape>? Shapes { get; set; }
		public List<Hyperlink>? Hyperlinks { get; set; }
	}
}
