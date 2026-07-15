using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Oqtane.Services;

namespace JBC.ExploreTheWorld.AL.Oqtane.CountrySlides
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

            // Reveal.js JS interop — scoped so each circuit gets its own module reference
            services.AddScoped<RevealJs__Interop__Interface, RevealJs__Interop>();
        }
    }
}
