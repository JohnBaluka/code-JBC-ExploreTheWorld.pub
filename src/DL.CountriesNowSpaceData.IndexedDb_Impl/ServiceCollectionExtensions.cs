using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.IndexedDb_Impl;
using JBC.ExploreTheWorld.DL;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.IndexedDb_Impl
{
    public static class ServiceCollectionExtensions
    {
        // Single-provider registration — binds the interface directly (for non-switching hosts).
        public static IServiceCollection AddExploreTheWorldIndexedDb(
            this IServiceCollection services)
        {
            services.AddScoped<IndexedDb_Repo__Interface, IndexedDb_Repo>();

            services.AddScoped<CountriesNowSpaceApiManager__Repo__Interface,
                               CountriesNowSpaceApiManager__IndexedDb__Repo>();

            return services;
        }

        // Keyed registration — participates in the runtime DbProviderSwitcher__Service.
        public static IServiceCollection AddExploreTheWorldIndexedDbProvider(
            this IServiceCollection services,
            string providerKey = DbProviderNames.IndexedDb)
        {
            services.AddScoped<IndexedDb_Repo__Interface, IndexedDb_Repo>();

            services.AddKeyedScoped<CountriesNowSpaceApiManager__Repo__Interface,
                                    CountriesNowSpaceApiManager__IndexedDb__Repo>(providerKey);

            return services;
        }
    }
}
