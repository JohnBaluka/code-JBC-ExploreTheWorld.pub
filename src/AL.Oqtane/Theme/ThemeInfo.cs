using System.Collections.Generic;
using global::Oqtane.Models;
using global::Oqtane.Shared;
using global::Oqtane.Themes;

namespace JBC.ExploreTheWorld.AL.Oqtane.Theme
{
    public class ThemeInfo : ITheme
    {
        public global::Oqtane.Models.Theme Theme => new global::Oqtane.Models.Theme
        {
            Name = "Explore the World Radzen Theme",
            Version = "1.0.0",
            ThemeSettingsType = "JBC.ExploreTheWorld.AL.Oqtane.Theme.ThemeSettings, JBC.ExploreTheWorld.AL.Oqtane.Theme",
            Resources = new List<Resource>
            {
                // Bootstrap is still required by the built-in Oqtane controls rendered by this
                // theme (Search, UserProfile, Login, ControlPanel offcanvas) and by the Oqtane
                // admin modules. The Radzen theme stylesheet is added in the layout <HeadContent>.
                new Stylesheet(Constants.BootstrapStylesheetUrl, Constants.BootstrapStylesheetIntegrity, "anonymous"),
                // Header/sidebar color sync — variable overrides, safe to load before the Radzen CSS
                new Stylesheet("_content/JBC.ExploreTheWorld.AL.Oqtane.Theme/css/theme.css"),
                new Script(Constants.BootstrapScriptUrl, Constants.BootstrapScriptIntegrity, "anonymous"),
                new Script("_content/Radzen.Blazor/Radzen.Blazor.js"),
                // Sidebar toggle: the theme renders statically, so the toggle is plain JS
                new Script("_content/JBC.ExploreTheWorld.AL.Oqtane.Theme/js/theme.js")
            }
        };
    }
}
