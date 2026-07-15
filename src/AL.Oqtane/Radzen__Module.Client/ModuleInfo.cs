using System.Collections.Generic;
using Oqtane.Models;
using Oqtane.Modules;

namespace JBC.ExploreTheWorld.AL.Oqtane.Radzen
{
    public class ModuleInfo : IModule
    {
        public ModuleDefinition ModuleDefinition => new ModuleDefinition
        {
            Name = "Radzen",
            Description = "Registers Radzen Blazor component suite and injects the Radzen default theme site-wide",
            Version = "1.0.0",
            ReleaseVersions = "1.0.0",
            Resources = new List<Resource>
            {
                new Script("_content/Radzen.Blazor/Radzen.Blazor.js"),
                new Stylesheet("_content/Radzen.Blazor/css/default.css")
            }
        };
    }
}
