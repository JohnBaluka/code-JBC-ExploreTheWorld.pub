using Blazored.SessionStorage;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SessionStorageDb_Impl;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SessionStorageDb_Impl
{
    public static class ServiceCollectionExtensions
    {
        // Single-provider registration — binds the interface directly (for non-switching hosts).
        public static IServiceCollection AddExploreTheWorldSessionStorageDb(
            this IServiceCollection services)
        {
            services.AddBlazoredSessionStorage();

            services.AddScoped<CountriesNowSpaceApiManager__Repo__Interface,
                               CountriesNowSpaceApiManager__SessionStorage__Repo>();

            return services;
        }

        // Keyed registration — participates in the runtime DbProviderSwitcher__Service.
        public static IServiceCollection AddExploreTheWorldSessionStorageDbProvider(
            this IServiceCollection services,
            string providerKey = DbProviderNames.SessionStorageDb)
        {
            services.AddBlazoredSessionStorage();

            services.AddKeyedScoped<CountriesNowSpaceApiManager__Repo__Interface,
                                    CountriesNowSpaceApiManager__SessionStorage__Repo>(providerKey);

            return services;
        }
    }
}
