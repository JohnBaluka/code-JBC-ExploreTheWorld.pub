using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word
{
    public class WordContentControlJs_Row
    {
        [JsonPropertyName("title")]
        public string? Title { get; set; }

        [JsonPropertyName("tag")]
        public string? Tag { get; set; }

        [JsonPropertyName("text")]
        public string? Text { get; set; }
    }
}
