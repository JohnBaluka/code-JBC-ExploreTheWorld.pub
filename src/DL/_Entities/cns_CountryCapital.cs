using System;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryCapital : cns_CountryCapital_Fields
    {
        [JsonIgnore]
        public cns_Country? Country { get; set; }
    }
}
