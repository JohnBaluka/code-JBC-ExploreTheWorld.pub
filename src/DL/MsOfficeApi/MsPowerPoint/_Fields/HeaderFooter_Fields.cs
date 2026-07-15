using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class HeaderFooter_Fields
    {
        public int? Format { get; set; }
        public string? Text { get; set; }
        public int? UseFormat { get; set; }
        public int? Visible { get; set; }
    }
}
