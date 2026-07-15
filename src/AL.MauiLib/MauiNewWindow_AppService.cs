namespace JBC.ExploreTheWorld.AL.MauiLib;

/// <summary>
/// Opens a new MAUI window sharing this app's DI container. Register only on
/// heads whose platform supports multi-window (WinUI, Mac Catalyst); other
/// heads register NullNewWindow_AppService instead.
/// </summary>
public class MauiNewWindow_AppService : NewWindow_AppService__Interface
{
    public void OpenNewWindow()
    {
        var app = Application.Current;
        if (app is null) return;
        app.Dispatcher.Dispatch(() =>
            app.OpenWindow(new Window(new MainPage()) { Title = App.AppTitle }));
    }
}
