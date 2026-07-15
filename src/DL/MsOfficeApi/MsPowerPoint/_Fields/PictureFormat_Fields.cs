using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class PictureFormat_Fields
    {
        public string? Brightness { get; set; }
        public int? ColorType { get; set; }
        public string? Contrast { get; set; }
        public string? CropBottom { get; set; }
        public string? CropLeft { get; set; }
        public string? CropRight { get; set; }
        public string? CropTop { get; set; }
        public long? TransparencyColor { get; set; }
        public int? TransparentBackground { get; set; }
    }
}
