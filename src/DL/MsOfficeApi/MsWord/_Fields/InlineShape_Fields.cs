using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class InlineShape_Fields
    {
        public int? Type { get; set; }
        public string? Width { get; set; }
        public string? Height { get; set; }
        public string? AlternativeText { get; set; }
    }
}
