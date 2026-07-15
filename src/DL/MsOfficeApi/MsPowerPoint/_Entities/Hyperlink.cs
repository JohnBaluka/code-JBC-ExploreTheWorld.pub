using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Hyperlink : Hyperlink_Fields
    {
        [JsonIgnore]
        public string CodeName
        {
            get
            {
                if (ShapeName == null) return string.Empty;

                return ShapeName.Replace(" ", "_").Replace("!", "_");
            }
        }
    }
}
