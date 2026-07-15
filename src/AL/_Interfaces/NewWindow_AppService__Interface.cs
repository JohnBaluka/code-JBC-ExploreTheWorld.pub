namespace JBC.ExploreTheWorld.AL
{
    /// <summary>
    /// Platform-neutral abstraction for opening a new application window or browser tab.
    /// Implemented per host (WinForms, Blazor browser, MAUI, or a no-op for task panes).
    /// </summary>
    public interface NewWindow_AppService__Interface
    {
        void OpenNewWindow();
    }
}
