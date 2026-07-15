using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl
{
    public static class ServiceCollectionExtensions
    {
        // Single-provider registration — binds the interface directly (for non-switching hosts).
        public static IServiceCollection AddExploreTheWorldSqliteDb(
            this IServiceCollection services,
            string connectionString)
        {
            services.AddDbContextFactory<ExploreTheWorldDbContext>(options =>
                options.UseSqlite(connectionString));

            services.AddTransient<CountriesNowSpaceApiManager__Repo__Interface,
                                  CountriesNowSpaceApiManager__Repo>();

            return services;
        }

        // Keyed registration — participates in the runtime DbProviderSwitcher__Service.
        // Registers a typed DbContext factory singleton plus a keyed repo under the provider key.
        // A null/empty connection string falls back to a %LocalAppData% SQLite file.
        public static IServiceCollection AddExploreTheWorldSqliteDbProvider(
            this IServiceCollection services,
            string? connectionString = null,
            string providerKey = DbProviderNames.SqliteDb)
        {
            if (string.IsNullOrEmpty(connectionString))
                connectionString = BuildDefaultSqliteConnectionString();

            var options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
                .UseSqlite(connectionString).Options;
            services.AddSingleton(new ExploreTheWorldDbContext__SqliteDb__Factory(options));

            services.AddKeyedScoped<CountriesNowSpaceApiManager__Repo__Interface>(providerKey, (sp, key) =>
                new CountriesNowSpaceApiManager__SqliteDb__Repo(
                    sp.GetRequiredService<ExploreTheWorldDbContext__SqliteDb__Factory>()));

            return services;
        }

        // Ensures the SQLite schema exists. Silently skipped if the provider is unavailable at startup.
        public static void TryEnsureExploreTheWorldSqliteDbCreated(this IServiceProvider services)
        {
            try
            {
                var factory = services.GetRequiredService<ExploreTheWorldDbContext__SqliteDb__Factory>();
                using var context = factory.CreateDbContext();
                context.Database.EnsureCreated();
            }
            catch { }
        }

        private static string BuildDefaultSqliteConnectionString()
        {
            var dbPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "JBC.ExploreTheWorld", "etw.db");
            Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);
            return $"Data Source={dbPath}";
        }
    }
}
