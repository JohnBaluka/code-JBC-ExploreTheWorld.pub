using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word
{
    // JS-shaped rows matching the object graph collected by getDocumentData() in
    // wwwroot/js/save-as-json.js (the Office.js Word object model). These are mapped
    // to the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord entities by MsWordJsMapper.
    public class WordDocumentJs_Row
    {
        [JsonPropertyName("properties")]
        public WordDocumentPropertiesJs_Row? Properties { get; set; }

        [JsonPropertyName("paragraphs")]
        public List<WordParagraphJs_Row> Paragraphs { get; set; } = new();

        [JsonPropertyName("tables")]
        public List<WordTableJs_Row> Tables { get; set; } = new();

        [JsonPropertyName("contentControls")]
        public List<WordContentControlJs_Row> ContentControls { get; set; } = new();

        [JsonPropertyName("inlinePictures")]
        public List<WordInlinePictureJs_Row> InlinePictures { get; set; } = new();
    }
}
