using System.Collections.Generic;
using Oqtane.Models;
using Oqtane.Modules;
using Oqtane.Shared;

namespace JBC.ExploreTheWorld.AL.Oqtane.CountriesNow
{
    public class ModuleInfo : IModule
    {
        public ModuleDefinition ModuleDefinition => new ModuleDefinition
        {
            Name = "CountriesNow",
            Description = "Displays countries and states sourced from the CountriesNow.space API, with SQL Server caching and browser-based Office export",
            Version = "1.0.0",
            ReleaseVersions = "1.0.0",
            Resources = new List<Resource>
            {
                new Stylesheet("_content/Radzen.Blazor/css/default.css"),
                new Script("_content/Radzen.Blazor/Radzen.Blazor.js")
                // download-file.js is now an ESM module loaded on demand by
                // FileDownload__Interop (DL.MsJSInterop); no global script needed.
            },
            PageTemplates = new List<PageTemplate>
            {
                new PageTemplate
                {
                    Name = "Countries Now",
                    Path = "countries-now",
                    Title = "Countries Now",
                    Order = 11,
                    IsNavigation = true,
                    IsClickable = true,
                    PermissionList = new List<Permission>
                    {
                        new Permission(PermissionNames.View, RoleNames.Everyone, true),
                        new Permission(PermissionNames.Edit, RoleNames.Admin, true)
                    },
                    PageTemplateModules = new List<PageTemplateModule>
                    {
                        new PageTemplateModule
                        {
                            Title = "Countries Now",
                            Pane = PaneNames.Default,
                            Order = 1,
                            PermissionList = new List<Permission>
                            {
                                new Permission(PermissionNames.View, RoleNames.Everyone, true),
                                new Permission(PermissionNames.Edit, RoleNames.Admin, true)
                            }
                        }
                    }
                }
            }
        };
    }
}
