using JBC.ExploreTheWorld.AL;
using JBC.ExploreTheWorld.AL.WinFormsLib;
using JBC.ExploreTheWorld.DL.MsOfficeApi_Impl;
using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;
using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Radzen;

namespace JBC.ExploreTheWorld.AL.WinFormApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            var config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true)
                .Build();

            var services = new ServiceCollection();
            services.AddWindowsFormsBlazorWebView();
#if DEBUG
            services.AddBlazorWebViewDeveloperTools();
#endif
            services.AddRadzenComponents();

            services.AddTransient<CountriesNowSpaceApi_Interface, CountriesNowSpaceApi__Repo>();

            // DbProvider_AppService — tracks the active provider; AvailableProviders drives the header dropdown.
            var dbProviderName = config["DbProvider"] ?? DbProviderNames.InMemoryDb;
            var dbProvider_AppService = new DbProvider_AppService
            {
                ProviderName       = dbProviderName,
                AvailableProviders = [DbProviderNames.InMemoryDb, DbProviderNames.AccessDb, DbProviderNames.SqliteDb, DbProviderNames.SqlServerDb],
            };
            services.AddSingleton(dbProvider_AppService);

            // Server DB providers — each registered as a keyed implementation; the switcher selects at call time.
            services.AddExploreTheWorldInMemoryDbProvider();
            services.AddExploreTheWorldAccessDbProvider(config.GetConnectionString("AccessDb"));
            services.AddExploreTheWorldSqliteDbProvider(config.GetConnectionString("SqliteDb"));
            services.AddExploreTheWorldSqlServerDbProvider(config.GetConnectionString("SqlServerDb"));

            // Switcher — reads DbProvider_AppService.ProviderName at call time to select the active implementation.
            services.AddCountriesNowSpaceDbSwitcher(() => dbProvider_AppService.ProviderName);

            services.AddTransient<CountriesNowSpaceManager__Service>();

            // Flag images — %LocalAppData% file cache with Wikimedia PNG download fallback.
            services.AddSingleton<FlagImageStore__Repo__Interface, FlagImageStore_FileSystem__Repo>();
            services.AddSingleton<FlagImageDownload__Repo__Interface, FlagImageDownload__Repo>();
            services.AddTransient<FlagImageManager__Service>();

            services.AddScoped<RevealJs__Interop__Interface, RevealJs__Interop>();

            // Generic JS interop (layout.js) — Main_Layout breakpoint badge
            services.AddScoped<Layout__Interop__Interface, Layout__Interop>();

            services.AddSingleton<WatcherEvent_AppService>();
            services.AddSingleton(new Layout_AppService { ShowSidebar = true, ShowWatcherNavItems = true, AppTitle = "ETW WinFormApp" });

            // Office export: BL manager + host repo factory (UI never instantiates DL repos directly).
            services.AddSingleton<MsOfficeExportRepoFactory__Interface, MsOfficeExportRepoFactory>();
            services.AddSingleton<MsOfficeExportManager__Service>();
            services.AddSingleton<OfficeExport_AppService__Interface, OfficeExport_AppService>();

            // NewWindow_AppService__Interface — opens a new ExploreTheWorld_Form sharing this DI container.
            services.AddSingleton<NewWindow_AppService__Interface, WinFormNewWindow_AppService>();

            var serviceProvider = services.BuildServiceProvider();

            // Ensure DB schema exists for every registered provider at startup.
            // Unavailable providers (SQL Server not installed, Access engine missing) are silently skipped.
            serviceProvider.TryEnsureExploreTheWorldInMemoryDbCreated();
            serviceProvider.TryEnsureExploreTheWorldAccessDbCreated();
            serviceProvider.TryEnsureExploreTheWorldSqliteDbCreated();
            serviceProvider.TryEnsureExploreTheWorldSqlServerDbCreated();

            // Supply the platform-specific Office composition (export factory + Save-As-JSON writer)
            // to the WinForms UI libraries, which reference no DL repo _Impl project.
            MsOfficeSaveAsJsonWriterProvider.Current  = new MsOfficeSaveAsJsonWriter();
            MsOfficeExportRepoFactoryProvider.Current = new MsOfficeExportRepoFactory();

            ApplicationConfiguration.Initialize();
            Application.Run(new ExploreTheWorld_Form(serviceProvider));
        }
    }
}
