using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Comment_Fields
    {
        public string? Author { get; set; }
        //public long? AuthorIndex { get; set; }
        public string? AuthorInitials { get; set; }
        public bool? Collapsed { get; set; }
        public string? DateTime { get; set; }
        public string? Left { get; set; }
        public string? ProviderID { get; set; }
        public string? Text { get; set; }
        //public long? TimeZoneBias { get; set; }
        public string? Top { get; set; }
        public string? UserID { get; set; }
    }
}
