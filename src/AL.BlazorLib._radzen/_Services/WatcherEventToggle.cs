using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.AL.BlazorLib
{
    /// <summary>
    /// Per-event toggle controlling whether that event is written to the watcher log.
    /// </summary>
    public sealed class WatcherEventToggle
    {
        public string Name     { get; }
        public string Category { get; }
        public bool   Log      { get; set; } = true;

        public WatcherEventToggle(string name, string category) { Name = name; Category = category; }
    }
}
