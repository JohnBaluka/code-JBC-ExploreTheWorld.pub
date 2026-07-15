using System;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_Country_Fields
    {
        public string Iso2 { get; set; } = string.Empty;
        public string? Country { get; set; }
        public string? Iso3 { get; set; }
    }
}
