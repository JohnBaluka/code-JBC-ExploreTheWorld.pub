using System;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_PopulationCount : cns_PopulationCount_Fields
    {
        [JsonIgnore]
        public cns_CountryPopulation? CountryPopulation { get; set; }
    }
}
