using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    public class CountryStates_Row
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("iso3")]
        public string? Iso3 { get; set; }

        [JsonPropertyName("states")]
        public List<CountryState_Row>? States { get; set; }
    }
}
