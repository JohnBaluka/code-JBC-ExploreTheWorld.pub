using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class LineFormat : LineFormat_Fields
    {
        public ColorFormat? BackColor { get; set; }
        public ColorFormat? ForeColor { get; set; }

    }
}
