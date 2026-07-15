using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl
{
    public static class ServiceCollectionExtensions
    {
        // Single-provider registration — binds the interface directly (for non-switching hosts).
        public static IServiceCollection AddExploreTheWorldAccessDb(
            this IServiceCollection services,
            string dbPath)
        {
            // OLE DB ACE connection string; requires Microsoft Access Database Engine installed.
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";

            services.AddDbContextFactory<ExploreTheWorldDbContext>(options =>
                options.UseJet(connectionString));

            services.AddTransient<CountriesNowSpaceApiManager__Repo__Interface,
                                  CountriesNowSpaceApiManager__Repo>();

            return services;
        }

        // Keyed registration — participates in the runtime DbProviderSwitcher__Service.
        // OLE DB ACE driver required; Windows-only.
#pragma warning disable CA1416
        public static IServiceCollection AddExploreTheWorldAccessDbProvider(
            this IServiceCollection services,
            string? dbPath = null,
            string providerKey = DbProviderNames.AccessDb)
        {
            if (string.IsNullOrEmpty(dbPath))
                dbPath = ExploreTheWorldAccessDb.DefaultDbPath;
            Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);

            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";
            var options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
                .UseJet(connectionString).Options;
            services.AddSingleton(new ExploreTheWorldDbContext__AccessDb__Factory(options));

            services.AddKeyedScoped<CountriesNowSpaceApiManager__Repo__Interface>(providerKey, (sp, key) =>
                new CountriesNowSpaceApiManager__AccessDb__Repo(
                    sp.GetRequiredService<ExploreTheWorldDbContext__AccessDb__Factory>()));

            return services;
        }
#pragma warning restore CA1416

        // Ensures the Access schema exists. Silently skipped if the ACE engine is missing.
        public static void TryEnsureExploreTheWorldAccessDbCreated(this IServiceProvider services)
        {
            try
            {
                var factory = services.GetRequiredService<ExploreTheWorldDbContext__AccessDb__Factory>();
                using var context = factory.CreateDbContext();
                context.Database.EnsureCreated();
            }
            catch { }
        }
    }
}
