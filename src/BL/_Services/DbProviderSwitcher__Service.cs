using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.BL
{
    // Generic runtime DB-provider switcher. Holds the set of registered provider implementations
    // (each keyed by its DbProviderNames key via keyed DI) and delegates every repo call to whichever
    // one is currently selected. The active provider name is read at call time via the injected
    // getActiveProviderName delegate (the host wires this to AL DbProvider_AppService.ProviderName), so
    // switching the selection in the UI takes effect immediately without restarting the app.
    //
    // Replaces the former DL ServerProxy/BrowserProxy repos: each host registers only the providers it
    // supports (server hosts add the EF Core _Impls; WASM hosts add the browser storage _Impls), and
    // this single service works for both.
    public sealed class DbProviderSwitcher__Service(
        IServiceProvider serviceProvider,
        Func<string> getActiveProviderName)
        : CountriesNowSpaceApiManager__Repo__Interface
    {
        private CountriesNowSpaceApiManager__Repo__Interface ActiveRepo =>
            serviceProvider.GetRequiredKeyedService<CountriesNowSpaceApiManager__Repo__Interface>(
                getActiveProviderName());

        // cns_Country
        public Task<List<cns_Country>> GetCountriesAsync() => ActiveRepo.GetCountriesAsync();
        public Task<cns_Country?> GetCountryByIso2Async(string iso2) => ActiveRepo.GetCountryByIso2Async(iso2);
        public Task<cns_Country> CreateCountryAsync(cns_Country entity) => ActiveRepo.CreateCountryAsync(entity);
        public Task<cns_Country> UpdateCountryAsync(cns_Country entity) => ActiveRepo.UpdateCountryAsync(entity);
        public Task<bool> DeleteCountryByIso2Async(string iso2) => ActiveRepo.DeleteCountryByIso2Async(iso2);

        // cns_City
        public Task<List<cns_City>> GetCitiesByIso2Async(string iso2) => ActiveRepo.GetCitiesByIso2Async(iso2);
        public Task<cns_City> CreateCityAsync(cns_City entity) => ActiveRepo.CreateCityAsync(entity);
        public Task<bool> DeleteCitiesByIso2Async(string iso2) => ActiveRepo.DeleteCitiesByIso2Async(iso2);

        // cns_CountryCapital
        public Task<cns_CountryCapital?> GetCountryCapitalByIso2Async(string iso2) => ActiveRepo.GetCountryCapitalByIso2Async(iso2);
        public Task<cns_CountryCapital> CreateCountryCapitalAsync(cns_CountryCapital entity) => ActiveRepo.CreateCountryCapitalAsync(entity);
        public Task<cns_CountryCapital> UpdateCountryCapitalAsync(cns_CountryCapital entity) => ActiveRepo.UpdateCountryCapitalAsync(entity);
        public Task<bool> DeleteCountryCapitalByIso2Async(string iso2) => ActiveRepo.DeleteCountryCapitalByIso2Async(iso2);

        // cns_CountryFlag
        public Task<List<cns_CountryFlag>> GetCountryFlagsAsync() => ActiveRepo.GetCountryFlagsAsync();
        public Task<cns_CountryFlag?> GetCountryFlagByIso2Async(string iso2) => ActiveRepo.GetCountryFlagByIso2Async(iso2);
        public Task<cns_CountryFlag> CreateCountryFlagAsync(cns_CountryFlag entity) => ActiveRepo.CreateCountryFlagAsync(entity);
        public Task<cns_CountryFlag> UpdateCountryFlagAsync(cns_CountryFlag entity) => ActiveRepo.UpdateCountryFlagAsync(entity);
        public Task<bool> DeleteCountryFlagByIso2Async(string iso2) => ActiveRepo.DeleteCountryFlagByIso2Async(iso2);

        // cns_CountryPopulation
        public Task<List<cns_CountryPopulation>> GetCountryPopulationsAsync() => ActiveRepo.GetCountryPopulationsAsync();
        public Task<cns_CountryPopulation?> GetCountryPopulationByIdAsync(int id) => ActiveRepo.GetCountryPopulationByIdAsync(id);
        public Task<cns_CountryPopulation> CreateCountryPopulationAsync(cns_CountryPopulation entity) => ActiveRepo.CreateCountryPopulationAsync(entity);
        public Task<bool> DeleteCountryPopulationByIdAsync(int id) => ActiveRepo.DeleteCountryPopulationByIdAsync(id);

        // cns_CountryStates
        public Task<List<cns_CountryStates>> GetCountryStatesAsync() => ActiveRepo.GetCountryStatesAsync();
        public Task<cns_CountryStates?> GetCountryStatesByIdAsync(int id) => ActiveRepo.GetCountryStatesByIdAsync(id);
        public Task<cns_CountryStates> CreateCountryStatesAsync(cns_CountryStates entity) => ActiveRepo.CreateCountryStatesAsync(entity);
        public Task<bool> DeleteCountryStatesByIdAsync(int id) => ActiveRepo.DeleteCountryStatesByIdAsync(id);

        // Bulk operations
        public Task ClearAllAsync() => ActiveRepo.ClearAllAsync();
    }
}
