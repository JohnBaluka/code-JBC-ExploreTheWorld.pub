using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    public class CountryPopulation_Row
    {
        [JsonPropertyName("country")]
        public string? Country { get; set; }

        [JsonPropertyName("code")]
        public string? Code { get; set; }

        [JsonPropertyName("iso3")]
        public string? Iso3 { get; set; }

        [JsonPropertyName("populationCounts")]
        public List<PopulationCount_Row>? PopulationCounts { get; set; }
    }
}
