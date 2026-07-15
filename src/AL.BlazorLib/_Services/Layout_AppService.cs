namespace JBC.ExploreTheWorld.AL.BlazorLib
{
    public class Layout_AppService
    {
        public bool ShowSidebar { get; set; } = true;
        public bool ShowWatcherNavItems { get; set; } = false;
        public bool ShowExportOptions { get; set; } = true;

        /// <summary>
        /// Text shown in the window/tab title bar. Each host sets this to its own
        /// project name (e.g. "ETW BlazorWebApp"); Main_Layout renders
        /// it as the page &lt;title&gt;. Defaults to the brand name when a host does not set it.
        /// </summary>
        public string AppTitle { get; set; } = "Explore the World";
    }
}
