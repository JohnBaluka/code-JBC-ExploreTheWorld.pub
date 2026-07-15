using System;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_PopulationCount_Fields
    {
        public int ID { get; set; }
        public int CountryPopulation_ID { get; set; }
        public string? Year { get; set; }
        public long Value { get; set; }
    }
}
