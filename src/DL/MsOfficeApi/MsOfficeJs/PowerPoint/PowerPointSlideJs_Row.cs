using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint
{
    public class PowerPointSlideJs_Row
    {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("index")]
        public int Index { get; set; }

        [JsonPropertyName("layoutName")]
        public string? LayoutName { get; set; }

        [JsonPropertyName("shapes")]
        public List<PowerPointShapeJs_Row> Shapes { get; set; } = new();
    }
}
