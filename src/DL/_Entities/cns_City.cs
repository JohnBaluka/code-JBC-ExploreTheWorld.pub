using System;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_City : cns_City_Fields
    {
        [JsonIgnore]
        public cns_Country? Country { get; set; }
    }
}
