using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl
{
    public static class ServiceCollectionExtensions
    {
        public const string DefaultDatabaseName = "JBC.ExploreTheWorld";

        // Single-provider registration — binds the interface directly (for non-switching hosts).
        public static IServiceCollection AddExploreTheWorldInMemoryDb(
            this IServiceCollection services,
            string databaseName = DefaultDatabaseName)
        {
            services.AddDbContextFactory<ExploreTheWorldDbContext>(options =>
                options.UseInMemoryDatabase(databaseName));

            services.AddTransient<CountriesNowSpaceApiManager__Repo__Interface,
                                  CountriesNowSpaceApiManager__Repo>();

            return services;
        }

        // Keyed registration — participates in the runtime DbProviderSwitcher__Service.
        public static IServiceCollection AddExploreTheWorldInMemoryDbProvider(
            this IServiceCollection services,
            string providerKey = DbProviderNames.InMemoryDb,
            string databaseName = DefaultDatabaseName)
        {
            var options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
                .UseInMemoryDatabase(databaseName).Options;
            services.AddSingleton(new ExploreTheWorldDbContext__InMemoryDb__Factory(options));

            services.AddKeyedScoped<CountriesNowSpaceApiManager__Repo__Interface>(providerKey, (sp, key) =>
                new CountriesNowSpaceApiManager__InMemoryDb__Repo(
                    sp.GetRequiredService<ExploreTheWorldDbContext__InMemoryDb__Factory>()));

            return services;
        }

        // Ensures the in-memory schema exists.
        public static void TryEnsureExploreTheWorldInMemoryDbCreated(this IServiceProvider services)
        {
            try
            {
                var factory = services.GetRequiredService<ExploreTheWorldDbContext__InMemoryDb__Factory>();
                using var context = factory.CreateDbContext();
                context.Database.EnsureCreated();
            }
            catch { }
        }
    }
}
