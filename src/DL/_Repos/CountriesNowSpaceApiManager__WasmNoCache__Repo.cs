using System.Collections.Generic;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL
{
    // WASM-compatible implementation of CountriesNowSpaceApiManager__Repo__Interface.
    // Always returns empty collections so the BL layer always falls through to the live API.
    // Write operations are no-ops (no database available in WebAssembly context).
    public class CountriesNowSpaceApiManager__WasmNoCache__Repo
        : CountriesNowSpaceApiManager__Repo__Interface
    {
        public Task<List<cns_Country>>          GetCountriesAsync()                             => Task.FromResult(new List<cns_Country>());
        public Task<cns_Country?>               GetCountryByIso2Async(string iso2)              => Task.FromResult<cns_Country?>(null);
        public Task<cns_Country>                CreateCountryAsync(cns_Country entity)          => Task.FromResult(entity);
        public Task<cns_Country>                UpdateCountryAsync(cns_Country entity)          => Task.FromResult(entity);
        public Task<bool>                       DeleteCountryByIso2Async(string iso2)           => Task.FromResult(true);

        public Task<List<cns_City>>             GetCitiesByIso2Async(string iso2)              => Task.FromResult(new List<cns_City>());
        public Task<cns_City>                   CreateCityAsync(cns_City entity)               => Task.FromResult(entity);
        public Task<bool>                       DeleteCitiesByIso2Async(string iso2)           => Task.FromResult(true);

        public Task<cns_CountryCapital?>        GetCountryCapitalByIso2Async(string iso2)      => Task.FromResult<cns_CountryCapital?>(null);
        public Task<cns_CountryCapital>         CreateCountryCapitalAsync(cns_CountryCapital e) => Task.FromResult(e);
        public Task<cns_CountryCapital>         UpdateCountryCapitalAsync(cns_CountryCapital e) => Task.FromResult(e);
        public Task<bool>                       DeleteCountryCapitalByIso2Async(string iso2)   => Task.FromResult(true);

        public Task<List<cns_CountryFlag>>      GetCountryFlagsAsync()                         => Task.FromResult(new List<cns_CountryFlag>());
        public Task<cns_CountryFlag?>           GetCountryFlagByIso2Async(string iso2)         => Task.FromResult<cns_CountryFlag?>(null);
        public Task<cns_CountryFlag>            CreateCountryFlagAsync(cns_CountryFlag e)      => Task.FromResult(e);
        public Task<cns_CountryFlag>            UpdateCountryFlagAsync(cns_CountryFlag e)      => Task.FromResult(e);
        public Task<bool>                       DeleteCountryFlagByIso2Async(string iso2)      => Task.FromResult(true);

        public Task<List<cns_CountryPopulation>> GetCountryPopulationsAsync()                  => Task.FromResult(new List<cns_CountryPopulation>());
        public Task<cns_CountryPopulation?>      GetCountryPopulationByIdAsync(int id)         => Task.FromResult<cns_CountryPopulation?>(null);
        public Task<cns_CountryPopulation>       CreateCountryPopulationAsync(cns_CountryPopulation e) => Task.FromResult(e);
        public Task<bool>                        DeleteCountryPopulationByIdAsync(int id)      => Task.FromResult(true);

        public Task<List<cns_CountryStates>>    GetCountryStatesAsync()                        => Task.FromResult(new List<cns_CountryStates>());
        public Task<cns_CountryStates?>         GetCountryStatesByIdAsync(int id)              => Task.FromResult<cns_CountryStates?>(null);
        public Task<cns_CountryStates>          CreateCountryStatesAsync(cns_CountryStates e)  => Task.FromResult(e);
        public Task<bool>                       DeleteCountryStatesByIdAsync(int id)           => Task.FromResult(true);

        public Task ClearAllAsync() => Task.CompletedTask;
    }
}
