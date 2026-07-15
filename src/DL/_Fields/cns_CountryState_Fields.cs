using System;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_CountryState_Fields
    {
        public int ID { get; set; }
        public int CountryStates_ID { get; set; }
        public string? Name { get; set; }
        public string? StateCode { get; set; }
    }
}
