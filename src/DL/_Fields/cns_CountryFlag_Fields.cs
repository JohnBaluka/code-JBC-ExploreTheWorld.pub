using System;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryFlag_Fields
    {
        public string Iso2 { get; set; } = string.Empty;
        public string? Name { get; set; }
        public string? Flag { get; set; }
        public string? DialCode { get; set; }
    }
}
