using System;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryPopulation_Fields
    {
        public int ID { get; set; }
        public string? Country { get; set; }
        public string? Code { get; set; }
        public string? Iso3 { get; set; }
    }
}
