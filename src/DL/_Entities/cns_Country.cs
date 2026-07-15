using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_Country : cns_Country_Fields
    {
        [JsonIgnore]
        public List<cns_City> Cities { get; set; } = new();

        [JsonIgnore]
        public cns_CountryCapital? Capital { get; set; }

        [JsonIgnore]
        public cns_CountryFlag? Flag { get; set; }
    }
}
