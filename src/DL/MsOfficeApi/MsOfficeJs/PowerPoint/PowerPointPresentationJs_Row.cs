using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint
{
    // JS-shaped rows matching the object graph collected by getPresentationData() in
    // wwwroot/js/save-as-json.js (the Office.js PowerPoint object model). These are
    // mapped to the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint entities by
    // MsPowerPointJsMapper.
    public class PowerPointPresentationJs_Row
    {
        [JsonPropertyName("title")]
        public string? Title { get; set; }

        [JsonPropertyName("slides")]
        public List<PowerPointSlideJs_Row> Slides { get; set; } = new();
    }
}
