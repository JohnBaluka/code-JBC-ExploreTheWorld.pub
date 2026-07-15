using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice
{
    [Serializable]
    public partial class ImageBlob_Fields
    {
        // File extension without the dot (e.g. "png").
        public string? Extension { get; set; }

        // Base64-encoded bytes when BlobOutput = Base64; otherwise null.
        public string? Base64 { get; set; }

        // Relative file reference when BlobOutput = SeparateFiles; otherwise null.
        public string? FileName { get; set; }
    }
}
