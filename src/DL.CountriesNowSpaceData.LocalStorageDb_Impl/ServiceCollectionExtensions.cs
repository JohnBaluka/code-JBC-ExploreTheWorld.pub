using Blazored.LocalStorage;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.LocalStorageDb_Impl;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.LocalStorageDb_Impl
{
    public static class ServiceCollectionExtensions
    {
        // Single-provider registration — binds the interface directly (for non-switching hosts).
        public static IServiceCollection AddExploreTheWorldLocalStorageDb(
            this IServiceCollection services)
        {
            services.AddBlazoredLocalStorage();

            services.AddScoped<CountriesNowSpaceApiManager__Repo__Interface,
                               CountriesNowSpaceApiManager__LocalStorage__Repo>();

            return services;
        }

        // Keyed registration — participates in the runtime DbProviderSwitcher__Service.
        public static IServiceCollection AddExploreTheWorldLocalStorageDbProvider(
            this IServiceCollection services,
            string providerKey = DbProviderNames.LocalStorageDb)
        {
            services.AddBlazoredLocalStorage();

            services.AddKeyedScoped<CountriesNowSpaceApiManager__Repo__Interface,
                                    CountriesNowSpaceApiManager__LocalStorage__Repo>(providerKey);

            return services;
        }
    }
}
