using System.Collections.Generic;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.AL.WinFormApp
{
    /// <summary>
    /// No-op DL manager for _netF: always returns empty collections so the BL layer always calls the API.
    /// </summary>
    internal class NullCountriesNowSpaceApiManager : CountriesNowSpaceApiManager__Repo__Interface
    {
        public Task<List<cns_Country>> GetCountriesAsync() => Task.FromResult(new List<cns_Country>());
        public Task<cns_Country?> GetCountryByIso2Async(string iso2) => Task.FromResult<cns_Country?>(null);
        public Task<cns_Country> CreateCountryAsync(cns_Country entity) => Task.FromResult(entity);
        public Task<cns_Country> UpdateCountryAsync(cns_Country entity) => Task.FromResult(entity);
        public Task<bool> DeleteCountryByIso2Async(string iso2) => Task.FromResult(true);

        public Task<List<cns_City>> GetCitiesByIso2Async(string iso2) => Task.FromResult(new List<cns_City>());
        public Task<cns_City> CreateCityAsync(cns_City entity) => Task.FromResult(entity);
        public Task<bool> DeleteCitiesByIso2Async(string iso2) => Task.FromResult(true);

        public Task<cns_CountryCapital?> GetCountryCapitalByIso2Async(string iso2) => Task.FromResult<cns_CountryCapital?>(null);
        public Task<cns_CountryCapital> CreateCountryCapitalAsync(cns_CountryCapital entity) => Task.FromResult(entity);
        public Task<cns_CountryCapital> UpdateCountryCapitalAsync(cns_CountryCapital entity) => Task.FromResult(entity);
        public Task<bool> DeleteCountryCapitalByIso2Async(string iso2) => Task.FromResult(true);

        public Task<List<cns_CountryFlag>> GetCountryFlagsAsync() => Task.FromResult(new List<cns_CountryFlag>());
        public Task<cns_CountryFlag?> GetCountryFlagByIso2Async(string iso2) => Task.FromResult<cns_CountryFlag?>(null);
        public Task<cns_CountryFlag> CreateCountryFlagAsync(cns_CountryFlag entity) => Task.FromResult(entity);
        public Task<cns_CountryFlag> UpdateCountryFlagAsync(cns_CountryFlag entity) => Task.FromResult(entity);
        public Task<bool> DeleteCountryFlagByIso2Async(string iso2) => Task.FromResult(true);

        public Task<List<cns_CountryPopulation>> GetCountryPopulationsAsync() => Task.FromResult(new List<cns_CountryPopulation>());
        public Task<cns_CountryPopulation?> GetCountryPopulationByIdAsync(int id) => Task.FromResult<cns_CountryPopulation?>(null);
        public Task<cns_CountryPopulation> CreateCountryPopulationAsync(cns_CountryPopulation entity) => Task.FromResult(entity);
        public Task<bool> DeleteCountryPopulationByIdAsync(int id) => Task.FromResult(true);

        public Task<List<cns_CountryStates>> GetCountryStatesAsync() => Task.FromResult(new List<cns_CountryStates>());
        public Task<cns_CountryStates?> GetCountryStatesByIdAsync(int id) => Task.FromResult<cns_CountryStates?>(null);
        public Task<cns_CountryStates> CreateCountryStatesAsync(cns_CountryStates entity) => Task.FromResult(entity);
        public Task<bool> DeleteCountryStatesByIdAsync(int id) => Task.FromResult(true);

        public Task ClearAllAsync() => Task.CompletedTask;
    }
}
