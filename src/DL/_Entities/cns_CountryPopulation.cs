using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryPopulation : cns_CountryPopulation_Fields
    {
        [JsonIgnore]
        public List<cns_PopulationCount> PopulationCounts { get; set; } = new();
    }
}
