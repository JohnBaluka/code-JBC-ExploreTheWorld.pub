using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Font_Fields
    {
		public int? AutoRotateNumbers { get; set; }
		public string? BaselineOffset { get; set; }
		public int? Bold { get; set; }
		//public int Embeddable { get; set; }
		//public int Embedded { get; set; }
		public int? Emboss { get; set; }
		public int? Italic { get; set; }
		public string? Name { get; set; }
		public string? NameAscii { get; set; }
		public string? NameComplexScript { get; set; }
		public string? NameFarEast { get; set; }
		public string? NameOther { get; set; }
		public int? Shadow { get; set; }
		public string? Size { get; set; }
		public int? Subscript { get; set; }
		public int? Superscript { get; set; }
		public int? Underline { get; set; }
	}
}
