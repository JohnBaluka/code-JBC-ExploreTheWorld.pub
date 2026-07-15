using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl
{
    public static class ServiceCollectionExtensions
    {
        private const string DefaultSqlServerConnectionString =
            "Server=.;Database=JBC_ExploreTheWorld;Trusted_Connection=True;TrustServerCertificate=True;";

        // Single-provider registration — binds the interface directly (for non-switching hosts).
        public static IServiceCollection AddExploreTheWorldSqlServerDb(
            this IServiceCollection services,
            string connectionString)
        {
            services.AddDbContextFactory<ExploreTheWorldDbContext>(options =>
                options.UseSqlServer(connectionString));

            services.AddTransient<CountriesNowSpaceApiManager__Repo__Interface,
                                  CountriesNowSpaceApiManager__Repo>();

            return services;
        }

        // Keyed registration — participates in the runtime DbProviderSwitcher__Service.
        public static IServiceCollection AddExploreTheWorldSqlServerDbProvider(
            this IServiceCollection services,
            string? connectionString = null,
            string providerKey = DbProviderNames.SqlServerDb)
        {
            if (string.IsNullOrEmpty(connectionString))
                connectionString = DefaultSqlServerConnectionString;

            var options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
                .UseSqlServer(connectionString).Options;
            services.AddSingleton(new ExploreTheWorldDbContext__SqlServerDb__Factory(options));

            services.AddKeyedScoped<CountriesNowSpaceApiManager__Repo__Interface>(providerKey, (sp, key) =>
                new CountriesNowSpaceApiManager__SqlServerDb__Repo(
                    sp.GetRequiredService<ExploreTheWorldDbContext__SqlServerDb__Factory>()));

            return services;
        }

        // Ensures the SQL Server schema exists. Silently skipped if SQL Server is unavailable at startup.
        public static void TryEnsureExploreTheWorldSqlServerDbCreated(this IServiceProvider services)
        {
            try
            {
                var factory = services.GetRequiredService<ExploreTheWorldDbContext__SqlServerDb__Factory>();
                using var context = factory.CreateDbContext();
                context.Database.EnsureCreated();
            }
            catch { }
        }
    }
}
