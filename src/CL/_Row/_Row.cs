using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.CL
{
    [Serializable]
    public class _Row
    {
        public virtual Guid GUID { get; set; }
        public virtual Double? Row_ID { get; set; }
        public virtual string Name { get; set; }
        public virtual string GetPrimaryKeyValue()
        {
            return GUID.ToString();
        }

        public virtual string GetBusinessKeyValue()
        {
            throw new NotImplementedException();
        }
    }
}
