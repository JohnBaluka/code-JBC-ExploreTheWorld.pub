using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryStates : cns_CountryStates_Fields
    {
        [JsonIgnore]
        public List<cns_CountryState> States { get; set; } = new();
    }
}
