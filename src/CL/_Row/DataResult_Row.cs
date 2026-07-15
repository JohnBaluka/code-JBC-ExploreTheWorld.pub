using System.Collections.Generic;

namespace JBC.ExploreTheWorld.CL
{
    /// <summary>
    /// Wraps a list of data items with the source from which they were loaded.
    /// </summary>
    public class DataResult_Row<T>
    {
        public List<T> Data { get; set; } = new List<T>();
        public DataSource_Enum Source { get; set; }
    }
}
