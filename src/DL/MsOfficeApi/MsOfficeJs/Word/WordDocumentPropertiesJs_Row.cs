using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word
{
    public class WordDocumentPropertiesJs_Row
    {
        [JsonPropertyName("title")]
        public string? Title { get; set; }

        [JsonPropertyName("subject")]
        public string? Subject { get; set; }

        [JsonPropertyName("author")]
        public string? Author { get; set; }

        [JsonPropertyName("keywords")]
        public string? Keywords { get; set; }

        [JsonPropertyName("comments")]
        public string? Comments { get; set; }

        [JsonPropertyName("lastAuthor")]
        public string? LastAuthor { get; set; }

        [JsonPropertyName("revisionNumber")]
        public string? RevisionNumber { get; set; }

        [JsonPropertyName("creationDate")]
        public string? CreationDate { get; set; }

        [JsonPropertyName("category")]
        public string? Category { get; set; }
    }
}
