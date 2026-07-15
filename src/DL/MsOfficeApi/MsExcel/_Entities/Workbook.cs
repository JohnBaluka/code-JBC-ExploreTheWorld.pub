using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class Workbook : Workbook_Fields
    {
        // Lists
        public List<BuiltInDocumentProperty>? BuiltInDocumentProperties { get; set; }
        public List<CustomDocumentProperty>? CustomDocumentProperties { get; set; }
        public List<Sheet>? Sheets { get; set; }
    }
}
