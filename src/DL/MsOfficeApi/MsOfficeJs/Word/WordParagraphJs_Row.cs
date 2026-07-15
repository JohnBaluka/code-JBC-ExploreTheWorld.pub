using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word
{
    public class WordParagraphJs_Row
    {
        [JsonPropertyName("text")]
        public string? Text { get; set; }

        [JsonPropertyName("style")]
        public string? Style { get; set; }

        // Office.js Word.Alignment value ("Left", "Centered", "Right", "Justified", ...).
        [JsonPropertyName("alignment")]
        public string? Alignment { get; set; }

        [JsonPropertyName("leftIndent")]
        public double? LeftIndent { get; set; }

        [JsonPropertyName("rightIndent")]
        public double? RightIndent { get; set; }

        [JsonPropertyName("firstLineIndent")]
        public double? FirstLineIndent { get; set; }

        [JsonPropertyName("spaceBefore")]
        public double? SpaceBefore { get; set; }

        [JsonPropertyName("spaceAfter")]
        public double? SpaceAfter { get; set; }

        [JsonPropertyName("lineSpacing")]
        public double? LineSpacing { get; set; }
    }
}
