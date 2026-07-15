using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word
{
    public class WordTableJs_Row
    {
        [JsonPropertyName("rowCount")]
        public long RowCount { get; set; }

        [JsonPropertyName("values")]
        public List<List<string?>>? Values { get; set; }
    }
}
