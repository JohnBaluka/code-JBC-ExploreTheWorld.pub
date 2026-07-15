namespace JBC.ExploreTheWorld.AL.MauiLib;

public class App : Application
{
    /// <summary>
    /// Text shown in the window title bar. Each platform head sets this to its own
    /// project name (e.g. "ETW MauiApp.WinUI") in its MauiProgram
    /// before the app is built. Defaults to the brand name when a head does not set it.
    /// </summary>
    public static string AppTitle { get; set; } = "Explore the World";

    protected override Window CreateWindow(IActivationState? activationState)
        => new Window(new MainPage()) { Title = AppTitle };
}
