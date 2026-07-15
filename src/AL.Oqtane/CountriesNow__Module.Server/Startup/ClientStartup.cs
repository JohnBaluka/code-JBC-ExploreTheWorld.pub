using JBC.ExploreTheWorld.AL;
using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.MsJSInterop.FileDownload;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Oqtane.Services;

namespace JBC.ExploreTheWorld.AL.Oqtane.CountriesNow
{
    public class ClientStartup : IClientStartup
    {
        public void ConfigureServices(IServiceCollection services)
        {
            // Radzen scoped services — also registered by Radzen__Module but safe to call again
            services.AddScoped<global::Radzen.DialogService>();
            services.AddScoped<global::Radzen.NotificationService>();
            services.AddScoped<global::Radzen.TooltipService>();
            services.AddScoped<global::Radzen.ContextMenuService>();

            services.AddDbContextFactory<ExploreTheWorldDbContext>((sp, options) =>
            {
                var config = sp.GetRequiredService<IConfiguration>();
                options.UseSqlServer(config.GetConnectionString("DefaultConnection")!);
            }, ServiceLifetime.Transient);

            services.AddTransient<CountriesNowSpaceApiManager__Repo__Interface, CountriesNowSpaceApiManager__Repo>();
            services.AddHttpClient<CountriesNowSpaceApi_Interface, CountriesNowSpaceApi__Repo>();
            services.AddTransient<CountriesNowSpaceManager__Service>();

            // Flag images — %LocalAppData% file cache with Wikimedia PNG download fallback.
            services.AddSingleton<FlagImageStore__Repo__Interface, FlagImageStore_FileSystem__Repo>();
            services.AddHttpClient<FlagImageDownload__Repo__Interface, FlagImageDownload__Repo>();
            services.AddTransient<FlagImageManager__Service>();

            // DbProvider_AppService — CountriesNow__Component requires it. The Oqtane host always
            // stores data in the Oqtane SQL Server database; AvailableProviders stays empty so
            // the provider dropdown is disabled (same convention as other server-side hosts).
            services.AddSingleton(new DbProvider_AppService { ProviderName = "SqlServerDb" });

            // Layout_AppService — CountriesNow__Component requires it; the defaults keep the
            // export panel visible (ShowExportOptions = true), matching the other export hosts.
            services.AddSingleton(new Layout_AppService());

            // Generic JS interop (download-file.js) required by BrowserExport_AppService
            services.AddScoped<FileDownload__Interop__Interface, FileDownload__Interop>();

            // Browser export: builds Office documents in-memory and triggers JS download
            services.AddScoped<OfficeExport_AppService__Interface, BrowserExport_AppService>();
        }
    }
}
