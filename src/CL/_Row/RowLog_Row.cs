using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.CL
{
    public class RowLog_Row
    {
        public Guid GUID { get; set; }
        public Double Row_ID { get; set; }
        public Guid RowLog_Row_GUID { get; set; }
        public Double Version_ID { get; set; }
        public DateTime DateTime { get; set; }
        public string Action { get; set; }
        public string UserName { get; set; }
        public string AppName { get; set; }
        public string NodeName { get; set; }
    }
}
