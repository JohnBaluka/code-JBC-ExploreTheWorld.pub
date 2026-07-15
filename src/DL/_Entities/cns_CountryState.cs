using System;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryState : cns_CountryState_Fields
    {
        [JsonIgnore]
        public cns_CountryStates? CountryStates { get; set; }
    }
}
