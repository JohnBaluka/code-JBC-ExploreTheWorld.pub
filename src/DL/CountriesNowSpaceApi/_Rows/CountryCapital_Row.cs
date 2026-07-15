using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    public class CountryCapital_Row
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("capital")]
        public string? Capital { get; set; }
    }
}
