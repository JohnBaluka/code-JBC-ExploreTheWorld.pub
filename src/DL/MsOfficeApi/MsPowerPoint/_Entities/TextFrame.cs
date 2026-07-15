using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class TextFrame : TextFrame_Fields
    {
        public TextRange? TextRange { get; set; }
    }
}
