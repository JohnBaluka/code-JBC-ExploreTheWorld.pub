using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.MsJSInterop.FileDownload;
using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs;
using JBC.ExploreTheWorld.AL.BlazorLib;
using Microsoft.Extensions.Logging;
using Radzen;

namespace JBC.ExploreTheWorld.AL.MauiLib;

/// <summary>
/// Shared MAUI bootstrap used by every platform head (Droid, iOS, Mac, WinUI).
/// Registers the platform-neutral services only; each head adds its own DB
/// providers, Layout_AppService flags, export/watcher/new-window services.
/// </summary>
public static class MauiProgramExtensions
{
    public static MauiAppBuilder UseSharedMauiLib(this MauiAppBuilder builder)
    {
        builder.UseMauiApp<App>();

        builder.Services.AddMauiBlazorWebView();
#if DEBUG
        builder.Services.AddBlazorWebViewDeveloperTools();
        builder.Logging.AddDebug();
#endif

        builder.Services.AddRadzenComponents();

        // BL orchestrators — each platform head supplies the concrete DL repos these depend on
        // (the CountriesNow API repo + flag-image store/download repos), so this shared library
        // references no DL repo `_Impl` project.
        builder.Services.AddTransient<CountriesNowSpaceManager__Service>();
        builder.Services.AddTransient<FlagImageManager__Service>();

        builder.Services.AddScoped<RevealJs__Interop__Interface, RevealJs__Interop>();

        // Generic JS interop (download-file.js, layout.js)
        builder.Services.AddScoped<FileDownload__Interop__Interface, FileDownload__Interop>();
        builder.Services.AddScoped<Layout__Interop__Interface, Layout__Interop>();

        // Idle by default — Watcher pages [Inject] it; connect buttons appear only
        // when a head (WinUI) assigns the Connect/Disconnect delegates.
        builder.Services.AddSingleton<WatcherEvent_AppService>();

        return builder;
    }
}
