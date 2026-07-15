using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.AL.MauiLib;
using JBC.ExploreTheWorld.AL.WinFormsLib;
using JBC.ExploreTheWorld.DL.MsOfficeApi_Impl;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;
using JBC.ExploreTheWorld.DL;
using Microsoft.Extensions.Configuration;

namespace JBC.ExploreTheWorld.AL.MauiApp.WinUI;

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
        global::JBC.ExploreTheWorld.AL.MauiLib.App.AppTitle = "ETW MauiApp.WinUI";

        // MAUI does not load appsettings.json automatically — pull it from the
        // unpackaged exe directory so DbProvider/ConnectionStrings apply.
        builder.Configuration.AddJsonFile(
            Path.Combine(AppContext.BaseDirectory, "appsettings.json"), optional: true);
        var config = builder.Configuration;

        // DbProvider_AppService — tracks the active provider; AvailableProviders drives the header dropdown.
        var dbProviderName = config["DbProvider"] ?? DbProviderNames.InMemoryDb;
        var dbProvider_AppService = new DbProvider_AppService
        {
            ProviderName       = dbProviderName,
            AvailableProviders = [DbProviderNames.InMemoryDb, DbProviderNames.AccessDb, DbProviderNames.SqliteDb, DbProviderNames.SqlServerDb],
        };
        builder.Services.AddSingleton(dbProvider_AppService);

        // Server DB providers — each registered as a keyed implementation; the switcher selects at call time.
        builder.Services.AddExploreTheWorldInMemoryDbProvider();
        builder.Services.AddExploreTheWorldAccessDbProvider(config.GetConnectionString("AccessDb"));
        builder.Services.AddExploreTheWorldSqliteDbProvider(config.GetConnectionString("SqliteDb"));
        builder.Services.AddExploreTheWorldSqlServerDbProvider(config.GetConnectionString("SqlServerDb"));

        // Switcher — reads DbProvider_AppService.ProviderName at call time to select the active implementation.
        builder.Services.AddCountriesNowSpaceDbSwitcher(() => dbProvider_AppService.ProviderName);

        builder.Services.AddSingleton<MauiWatcher_AppService>();
        builder.Services.AddSingleton(new Layout_AppService
        {
            ShowSidebar         = true,
            ShowWatcherNavItems = true,
        });

        // Supply the platform-specific Save-As-JSON writer to the shared watcher code
        // (MauiWatcher_AppService), which reaches it via MsOfficeSaveAsJsonWriterProvider.
        MsOfficeSaveAsJsonWriterProvider.Current = new MsOfficeSaveAsJsonWriter();

        // Office export: BL manager + host repo factory (UI never instantiates DL repos directly).
        // Windows head uses the real Office export pipeline (Interop/Dynamic/NetOffice/OpenXML),
        // not the browser OpenXML-only fallback.
        builder.Services.AddSingleton<MsOfficeExportRepoFactory__Interface, MsOfficeExportRepoFactory>();
        builder.Services.AddSingleton<MsOfficeExportManager__Service>();
        builder.Services.AddSingleton<OfficeExport_AppService__Interface, OfficeExport_AppService>();

        // NewWindow_AppService__Interface — opens a new MAUI window sharing this app's DI container.
        builder.Services.AddSingleton<NewWindow_AppService__Interface, MauiNewWindow_AppService>();

        var app = builder.Build();

        // Ensure DB schema exists for every registered provider at startup.
        // Unavailable providers (SQL Server not installed, Access engine missing) are silently skipped.
        app.Services.TryEnsureExploreTheWorldInMemoryDbCreated();
        app.Services.TryEnsureExploreTheWorldAccessDbCreated();
        app.Services.TryEnsureExploreTheWorldSqliteDbCreated();
        app.Services.TryEnsureExploreTheWorldSqlServerDbCreated();
        // Eagerly instantiate so Word/Excel/PowerPoint connect delegates are registered
        // on the singleton WatcherEvent_AppService before any Blazor component resolves it.
        app.Services.GetRequiredService<MauiWatcher_AppService>();

        return app;
    }
}
