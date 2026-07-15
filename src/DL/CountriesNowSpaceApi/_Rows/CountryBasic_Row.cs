using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    public class CountryBasic_Row
    {
        [JsonPropertyName("country")]
        public string? Country { get; set; }

        [JsonPropertyName("iso2")]
        public string? Iso2 { get; set; }

        [JsonPropertyName("iso3")]
        public string? Iso3 { get; set; }

        [JsonPropertyName("cities")]
        public List<string>? Cities { get; set; }
    }
}
