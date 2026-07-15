using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    public class CountryState_Row
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("state_code")]
        public string? StateCode { get; set; }
    }
}
