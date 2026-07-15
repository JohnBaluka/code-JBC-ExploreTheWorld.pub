using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.AL.MauiLib;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;

namespace JBC.ExploreTheWorld.AL.MauiApp.Mac;

public static class MauiProgram
{
    public static global::Microsoft.Maui.Hosting.MauiApp CreateMauiApp()
    {
        var builder = global::Microsoft.Maui.Hosting.MauiApp.CreateBuilder();
        builder.UseSharedMauiLib();

        // Concrete DL repos supplied by this head (the shared MauiLib references no DL repo _Impl).
        builder.Services.AddTransient<CountriesNowSpaceApi_Interface, CountriesNowSpaceApi__Repo>();
        builder.Services.AddSingleton<FlagImageStore__Repo__Interface, FlagImageStore_FileSystem__Repo>();
        builder.Services.AddSingleton<FlagImageDownload__Repo__Interface, FlagImageDownload__Repo>();

        // Window title bar shows this head's project name.
        global::JBC.ExploreTheWorld.AL.MauiLib.App.AppTitle = "ETW MauiApp.Mac";

        // InMemoryDb is the only DB provider supported on this platform.
        builder.Services.AddSingleton(new DbProvider_AppService
        {
            ProviderName       = "InMemoryDb",
            AvailableProviders = ["InMemoryDb"],
        });
        builder.Services.AddExploreTheWorldInMemoryDb();

        // Watchers and Export API Data are WinUI-only features.
        builder.Services.AddSingleton(new Layout_AppService
        {
            ShowSidebar         = true,
            ShowWatcherNavItems = false,
            ShowExportOptions   = false,
        });

        // Mac Catalyst supports MAUI multi-window — reuse the shared MAUI implementation.
        builder.Services.AddSingleton<NewWindow_AppService__Interface, MauiNewWindow_AppService>();

        return builder.Build();
    }
}
