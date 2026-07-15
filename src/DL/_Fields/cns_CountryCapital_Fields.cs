using System;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryCapital_Fields
    {
        public string Iso2 { get; set; } = string.Empty;
        public string? Name { get; set; }
        public string? Capital { get; set; }
    }
}
