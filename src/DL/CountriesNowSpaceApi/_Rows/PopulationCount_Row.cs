using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    public class PopulationCount_Row
    {
        [JsonPropertyName("year")]
        public string? Year { get; set; }

        [JsonPropertyName("value")]
        public long Value { get; set; }
    }
}
