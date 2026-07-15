using System.Diagnostics;
using JBC.ExploreTheWorld.DL.MsOfficeApi_Impl;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.AL.ExportData.ConsoleApp
{
    /// <summary>
    /// Console version of the Countries API form's export: loads the countries through the
    /// DB-backed manager (check DB → API fallback → persist), enriches them with cached flag
    /// images, and exports to Word/Excel/PowerPoint with the selected method.
    /// </summary>
    internal static class Program
    {
        private static async Task<int> Main(string[] args)
        {
            var options = ExportDataOptions.Parse(args);

            if (options.ShowHelp)
            {
                Console.WriteLine(ExportDataOptions.Usage);
                return 0;
            }

            if (options.Error != null)
            {
                Console.Error.WriteLine(options.Error);
                Console.Error.WriteLine();
                Console.Error.WriteLine(ExportDataOptions.Usage);
                return 1;
            }

            try
            {
                using var serviceProvider = BuildServiceProvider(options.Provider);
                serviceProvider.TryEnsureExploreTheWorldSqliteDbCreated();
                serviceProvider.TryEnsureExploreTheWorldSqlServerDbCreated();
                serviceProvider.TryEnsureExploreTheWorldInMemoryDbCreated();
                serviceProvider.TryEnsureExploreTheWorldAccessDbCreated();

                Log($"Loading countries ({options.Provider})...");
                var countriesService = serviceProvider.GetRequiredService<CountriesNowSpaceManager__Service>();
                var result = await countriesService.GetAllCountriesAsync();
                Log($"✔ Loaded {result.Data.Count} countries (source: {result.Source}).");

                var rows = result.Data
                    .Select(c => new MsOfficeCountry_Row(c.Country, c.Iso2, c.Iso3))
                    .ToList();

                Log($"Export target : {options.FilePath}");
                Log($"Export type   : {options.Type} — {options.Method}");

                var stopwatch = Stopwatch.StartNew();
                var exportManager = serviceProvider.GetRequiredService<MsOfficeExportManager__Service>();
                await exportManager.ExportAsync(rows, options.Type, options.Method, options.FilePath, Log);
                stopwatch.Stop();

                Log($"✔ Export completed successfully in {Duration_Helper.Format(stopwatch.Elapsed)}.");
                return 0;
            }
            catch (Exception ex)
            {
                Log($"✘ Export failed: {ex.Message}");
                return 1;
            }
        }

        // Mirrors the AL.WinFormApp service wiring for the countries/export pipeline.
        private static ServiceProvider BuildServiceProvider(string providerName)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: true)
                .Build();

            var services = new ServiceCollection();

            services.AddTransient<CountriesNowSpaceApi_Interface, CountriesNowSpaceApi__Repo>();

            services.AddExploreTheWorldSqliteDbProvider(config.GetConnectionString("SqliteDb"));
            services.AddExploreTheWorldSqlServerDbProvider(config.GetConnectionString("SqlServerDb"));
            services.AddExploreTheWorldInMemoryDbProvider();
            services.AddExploreTheWorldAccessDbProvider(config.GetConnectionString("AccessDb"));

            services.AddCountriesNowSpaceDbSwitcher(() => providerName);

            services.AddTransient<CountriesNowSpaceManager__Service>();

            // Flag images — %LocalAppData% file cache with Wikimedia PNG download fallback.
            services.AddSingleton<FlagImageStore__Repo__Interface, FlagImageStore_FileSystem__Repo>();
            services.AddSingleton<FlagImageDownload__Repo__Interface, FlagImageDownload__Repo>();
            services.AddTransient<FlagImageManager__Service>();

            // Office export: BL manager + host repo factory (the factory is the single place
            // that instantiates the platform-specific DL repos).
            services.AddSingleton<MsOfficeExportRepoFactory__Interface, MsOfficeExportRepoFactory>();
            services.AddSingleton<MsOfficeExportManager__Service>();

            return services.BuildServiceProvider();
        }

        private static void Log(string message)
            => Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
    }
}
