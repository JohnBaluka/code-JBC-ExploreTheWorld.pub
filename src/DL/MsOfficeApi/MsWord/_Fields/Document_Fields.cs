using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
{
    [Serializable]
    public partial class Document_Fields
    {
        public string? Name { get; set; }
        public string? FullName { get; set; }
        public string? Path { get; set; }
        public bool? Saved { get; set; }
        public int? SaveFormat { get; set; }
        public bool? ReadOnly { get; set; }
        public int? ProtectionType { get; set; }
        public bool? TrackRevisions { get; set; }
        public bool? VBASigned { get; set; }
    }
}
