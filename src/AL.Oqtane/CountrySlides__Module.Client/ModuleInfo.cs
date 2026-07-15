using System.Collections.Generic;
using Oqtane.Models;
using Oqtane.Modules;
using Oqtane.Shared;

namespace JBC.ExploreTheWorld.AL.Oqtane.CountrySlides
{
    public class ModuleInfo : IModule
    {
        public ModuleDefinition ModuleDefinition => new ModuleDefinition
        {
            Name = "CountrySlides",
            Description = "Reveal.js slideshow presenting countries from the CountriesNow.space API with interactive controls",
            Version = "1.0.0",
            ReleaseVersions = "1.0.0",
            Resources = new List<Resource>
            {
                new Stylesheet("_content/Radzen.Blazor/css/default.css"),
                // etw-slide-* layout styles used by CountrySlides__Component
                // (slides-only file — app.css would impose app-shell globals on the site)
                new Stylesheet("_content/JBC.ExploreTheWorld.AL.BlazorLib/css/country-slides.css"),
                new Script("_content/Radzen.Blazor/Radzen.Blazor.js")
            },
            PageTemplates = new List<PageTemplate>
            {
                new PageTemplate
                {
                    Name = "Country Slides",
                    Path = "country-slides",
                    Title = "Country Slides",
                    Order = 12,
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
                            Title = "Country Slides",
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
