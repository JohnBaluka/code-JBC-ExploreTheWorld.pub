using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
{
    [Serializable]
    public partial class Cell : Cell_Fields
    {
        // Entities
        public Font? Font { get; set; }
        public Interior? Interior { get; set; }
    }
}
