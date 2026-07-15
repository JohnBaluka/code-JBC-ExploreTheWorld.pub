using System;
using JBC.ExploreTheWorld.DL;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.BL
{
    public static class ServiceCollectionExtensions
    {
        // Registers the runtime DB-provider switcher as the CountriesNowSpaceApiManager__Repo__Interface.
        // The host must first register one or more keyed provider implementations (AddExploreTheWorld*DbProvider),
        // then call this passing a delegate that returns the currently-selected provider key
        // (typically () => dbProvider_AppService.ProviderName).
        public static IServiceCollection AddCountriesNowSpaceDbSwitcher(
            this IServiceCollection services,
            Func<string> getActiveProviderName)
        {
            services.AddScoped<CountriesNowSpaceApiManager__Repo__Interface>(sp =>
                new DbProviderSwitcher__Service(sp, getActiveProviderName));

            return services;
        }
    }
}
