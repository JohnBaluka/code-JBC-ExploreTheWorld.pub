using Blazored.LocalStorage;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.LocalStorageDb_Impl
{
    // Browser localStorage implementation of CountriesNowSpaceApiManager__Repo__Interface.
    // Entities are serialized as JSON lists keyed by constant storage keys.
    // All operations are async; read methods return empty when the key has not been written yet.
    public class CountriesNowSpaceApiManager__LocalStorage__Repo
        : CountriesNowSpaceApiManager__Repo__Interface
    {
        private readonly ILocalStorageService _storage;

        private const string KeyCountries   = "etw.cns.countries";
        private const string KeyCities      = "etw.cns.cities";
        private const string KeyCapitals    = "etw.cns.capitals";
        private const string KeyFlags       = "etw.cns.flags";
        private const string KeyPopulations = "etw.cns.populations";
        private const string KeyStates      = "etw.cns.states";

        public CountriesNowSpaceApiManager__LocalStorage__Repo(ILocalStorageService storage)
        {
            _storage = storage;
        }

        // ================================================================
        // cns_Country
        // ================================================================

        public async Task<List<cns_Country>> GetCountriesAsync()
            => await _storage.GetItemAsync<List<cns_Country>>(KeyCountries) ?? [];

        public async Task<cns_Country?> GetCountryByIso2Async(string iso2)
        {
            var list = await GetCountriesAsync();
            return list.FirstOrDefault(e => e.Iso2 == iso2);
        }

        public async Task<cns_Country> CreateCountryAsync(cns_Country entity)
        {
            var list = await GetCountriesAsync();
            list.RemoveAll(e => e.Iso2 == entity.Iso2);
            list.Add(entity);
            await _storage.SetItemAsync(KeyCountries, list);
            return entity;
        }

        public async Task<cns_Country> UpdateCountryAsync(cns_Country entity)
            => await CreateCountryAsync(entity);

        public async Task<bool> DeleteCountryByIso2Async(string iso2)
        {
            var list = await GetCountriesAsync();
            var removed = list.RemoveAll(e => e.Iso2 == iso2);
            if (removed > 0) await _storage.SetItemAsync(KeyCountries, list);
            return removed > 0;
        }

        // ================================================================
        // cns_City
        // ================================================================

        public async Task<List<cns_City>> GetCitiesByIso2Async(string iso2)
        {
            var list = await _storage.GetItemAsync<List<cns_City>>(KeyCities) ?? [];
            return list.Where(e => e.Iso2 == iso2).ToList();
        }

        public async Task<cns_City> CreateCityAsync(cns_City entity)
        {
            var list = await _storage.GetItemAsync<List<cns_City>>(KeyCities) ?? [];
            list.Add(entity);
            await _storage.SetItemAsync(KeyCities, list);
            return entity;
        }

        public async Task<bool> DeleteCitiesByIso2Async(string iso2)
        {
            var list = await _storage.GetItemAsync<List<cns_City>>(KeyCities) ?? [];
            var removed = list.RemoveAll(e => e.Iso2 == iso2);
            if (removed > 0) await _storage.SetItemAsync(KeyCities, list);
            return removed > 0;
        }

        // ================================================================
        // cns_CountryCapital
        // ================================================================

        public async Task<cns_CountryCapital?> GetCountryCapitalByIso2Async(string iso2)
        {
            var list = await _storage.GetItemAsync<List<cns_CountryCapital>>(KeyCapitals) ?? [];
            return list.FirstOrDefault(e => e.Iso2 == iso2);
        }

        public async Task<cns_CountryCapital> CreateCountryCapitalAsync(cns_CountryCapital entity)
        {
            var list = await _storage.GetItemAsync<List<cns_CountryCapital>>(KeyCapitals) ?? [];
            list.RemoveAll(e => e.Iso2 == entity.Iso2);
            list.Add(entity);
            await _storage.SetItemAsync(KeyCapitals, list);
            return entity;
        }

        public async Task<cns_CountryCapital> UpdateCountryCapitalAsync(cns_CountryCapital entity)
            => await CreateCountryCapitalAsync(entity);

        public async Task<bool> DeleteCountryCapitalByIso2Async(string iso2)
        {
            var list = await _storage.GetItemAsync<List<cns_CountryCapital>>(KeyCapitals) ?? [];
            var removed = list.RemoveAll(e => e.Iso2 == iso2);
            if (removed > 0) await _storage.SetItemAsync(KeyCapitals, list);
            return removed > 0;
        }

        // ================================================================
        // cns_CountryFlag
        // ================================================================

        public async Task<List<cns_CountryFlag>> GetCountryFlagsAsync()
            => await _storage.GetItemAsync<List<cns_CountryFlag>>(KeyFlags) ?? [];

        public async Task<cns_CountryFlag?> GetCountryFlagByIso2Async(string iso2)
        {
            var list = await GetCountryFlagsAsync();
            return list.FirstOrDefault(e => e.Iso2 == iso2);
        }

        public async Task<cns_CountryFlag> CreateCountryFlagAsync(cns_CountryFlag entity)
        {
            var list = await GetCountryFlagsAsync();
            list.RemoveAll(e => e.Iso2 == entity.Iso2);
            list.Add(entity);
            await _storage.SetItemAsync(KeyFlags, list);
            return entity;
        }

        public async Task<cns_CountryFlag> UpdateCountryFlagAsync(cns_CountryFlag entity)
            => await CreateCountryFlagAsync(entity);

        public async Task<bool> DeleteCountryFlagByIso2Async(string iso2)
        {
            var list = await GetCountryFlagsAsync();
            var removed = list.RemoveAll(e => e.Iso2 == iso2);
            if (removed > 0) await _storage.SetItemAsync(KeyFlags, list);
            return removed > 0;
        }

        // ================================================================
        // cns_CountryPopulation
        // ================================================================

        public async Task<List<cns_CountryPopulation>> GetCountryPopulationsAsync()
            => await _storage.GetItemAsync<List<cns_CountryPopulation>>(KeyPopulations) ?? [];

        public async Task<cns_CountryPopulation?> GetCountryPopulationByIdAsync(int id)
        {
            var list = await GetCountryPopulationsAsync();
            return list.FirstOrDefault(e => e.ID == id);
        }

        public async Task<cns_CountryPopulation> CreateCountryPopulationAsync(cns_CountryPopulation entity)
        {
            var list = await GetCountryPopulationsAsync();
            list.Add(entity);
            await _storage.SetItemAsync(KeyPopulations, list);
            return entity;
        }

        public async Task<bool> DeleteCountryPopulationByIdAsync(int id)
        {
            var list = await GetCountryPopulationsAsync();
            var removed = list.RemoveAll(e => e.ID == id);
            if (removed > 0) await _storage.SetItemAsync(KeyPopulations, list);
            return removed > 0;
        }

        // ================================================================
        // cns_CountryStates
        // ================================================================

        public async Task<List<cns_CountryStates>> GetCountryStatesAsync()
            => await _storage.GetItemAsync<List<cns_CountryStates>>(KeyStates) ?? [];

        public async Task<cns_CountryStates?> GetCountryStatesByIdAsync(int id)
        {
            var list = await GetCountryStatesAsync();
            return list.FirstOrDefault(e => e.ID == id);
        }

        public async Task<cns_CountryStates> CreateCountryStatesAsync(cns_CountryStates entity)
        {
            var list = await GetCountryStatesAsync();
            list.Add(entity);
            await _storage.SetItemAsync(KeyStates, list);
            return entity;
        }

        public async Task<bool> DeleteCountryStatesByIdAsync(int id)
        {
            var list = await GetCountryStatesAsync();
            var removed = list.RemoveAll(e => e.ID == id);
            if (removed > 0) await _storage.SetItemAsync(KeyStates, list);
            return removed > 0;
        }

        // ================================================================
        // Bulk operations
        // ================================================================

        public async Task ClearAllAsync()
        {
            await _storage.RemoveItemAsync(KeyCountries);
            await _storage.RemoveItemAsync(KeyCities);
            await _storage.RemoveItemAsync(KeyCapitals);
            await _storage.RemoveItemAsync(KeyFlags);
            await _storage.RemoveItemAsync(KeyPopulations);
            await _storage.RemoveItemAsync(KeyStates);
        }
    }
}
