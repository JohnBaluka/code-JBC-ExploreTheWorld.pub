using System;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryFlag : cns_CountryFlag_Fields
    {
        [JsonIgnore]
        public cns_Country? Country { get; set; }
    }
}
