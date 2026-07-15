using JBC.ExploreTheWorld.AL.BlazorLib;
using Microsoft.AspNetCore.Components.WebView.Maui;

namespace JBC.ExploreTheWorld.AL.MauiLib;

public class MainPage : ContentPage
{
    public MainPage()
    {
        Title = App.AppTitle;
        var blazorWebView = new BlazorWebView { HostPage = "wwwroot/index.html" };
        blazorWebView.RootComponents.Add(new RootComponent
        {
            Selector = "#app",
            ComponentType = typeof(Routes)
        });
        blazorWebView.RootComponents.Add(new RootComponent
        {
            Selector = "head::after",
            ComponentType = typeof(Microsoft.AspNetCore.Components.Web.HeadOutlet)
        });
        Content = blazorWebView;
    }
}
