using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint
{
    public class PowerPointShapeJs_Row
    {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        // Office.js PowerPoint.ShapeType value (e.g. "GeometricShape", "Image", "Group").
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("left")]
        public double? Left { get; set; }

        [JsonPropertyName("top")]
        public double? Top { get; set; }

        [JsonPropertyName("width")]
        public double? Width { get; set; }

        [JsonPropertyName("height")]
        public double? Height { get; set; }

        [JsonPropertyName("altTextDescription")]
        public string? AltTextDescription { get; set; }

        [JsonPropertyName("altTextTitle")]
        public string? AltTextTitle { get; set; }

        [JsonPropertyName("zOrderIndex")]
        public int ZOrderIndex { get; set; }

        [JsonPropertyName("text")]
        public string? Text { get; set; }
    }
}
