using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word
{
    public class WordInlinePictureJs_Row
    {
        [JsonPropertyName("base64")]
        public string? Base64 { get; set; }

        [JsonPropertyName("width")]
        public double? Width { get; set; }

        [JsonPropertyName("height")]
        public double? Height { get; set; }

        [JsonPropertyName("altTextDescription")]
        public string? AltTextDescription { get; set; }
    }
}
