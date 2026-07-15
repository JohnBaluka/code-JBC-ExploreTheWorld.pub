using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class CustomDocumentProperty_Fields
    {
        public string? Name { get; set; }
        public long? Creator { get; set; }
        public string? LinkSource { get; set; }
        public int? LinkToContent { get; set; }
        public int? Type { get; set; }
        public string? Value { get; set; }

    }
}
