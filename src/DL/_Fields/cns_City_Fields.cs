using System;

namespace JBC.ExploreTheWorld.DL
{
    [Serializable]
    public partial class cns_City_Fields
    {
        public int ID { get; set; }
        public string Iso2 { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
    }
}
