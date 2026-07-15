using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class Workbook_Fields
    {
        public string? Name { get; set; }
        public string? FullName { get; set; }
        public string? Path { get; set; }
        public bool? Saved { get; set; }
        public bool? ReadOnly { get; set; }
        public bool? HasVBProject { get; set; }
        public int? FileFormat { get; set; }
        public string? CodeName { get; set; }
    }
}
