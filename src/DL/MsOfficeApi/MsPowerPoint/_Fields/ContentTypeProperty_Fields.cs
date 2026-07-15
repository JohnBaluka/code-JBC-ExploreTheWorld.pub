using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class ContentTypeProperty_Fields
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public long? Creator { get; set; }
        public bool? IsReadOnly { get; set; }
        public bool? IsRequired { get; set; }
        public int? Type { get; set; }
        public string? Value { get; set; }

    }
}
