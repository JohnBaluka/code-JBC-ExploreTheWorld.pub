using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class HeadersFooters : HeadersFooters_Fields
    {
        // Entities
        public HeaderFooter? DateAndTime { get; set; }
        public HeaderFooter? Footer { get; set; }
        public HeaderFooter? Header { get; set; }
        public HeaderFooter? SlideNumber { get; set; }
    }
}
