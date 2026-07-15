using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL
{
    public interface CountriesNowSpaceApiManager__Repo__Interface
    {
        // cns_Country
        Task<List<cns_Country>> GetCountriesAsync();
        Task<cns_Country?> GetCountryByIso2Async(string iso2);
        Task<cns_Country> CreateCountryAsync(cns_Country entity);
        Task<cns_Country> UpdateCountryAsync(cns_Country entity);
        Task<bool> DeleteCountryByIso2Async(string iso2);

        // cns_City
        Task<List<cns_City>> GetCitiesByIso2Async(string iso2);
        Task<cns_City> CreateCityAsync(cns_City entity);
        Task<bool> DeleteCitiesByIso2Async(string iso2);

        // cns_CountryCapital
        Task<cns_CountryCapital?> GetCountryCapitalByIso2Async(string iso2);
        Task<cns_CountryCapital> CreateCountryCapitalAsync(cns_CountryCapital entity);
        Task<cns_CountryCapital> UpdateCountryCapitalAsync(cns_CountryCapital entity);
        Task<bool> DeleteCountryCapitalByIso2Async(string iso2);

        // cns_CountryFlag
        Task<List<cns_CountryFlag>> GetCountryFlagsAsync();
        Task<cns_CountryFlag?> GetCountryFlagByIso2Async(string iso2);
        Task<cns_CountryFlag> CreateCountryFlagAsync(cns_CountryFlag entity);
        Task<cns_CountryFlag> UpdateCountryFlagAsync(cns_CountryFlag entity);
        Task<bool> DeleteCountryFlagByIso2Async(string iso2);

        // cns_CountryPopulation
        Task<List<cns_CountryPopulation>> GetCountryPopulationsAsync();
        Task<cns_CountryPopulation?> GetCountryPopulationByIdAsync(int id);
        Task<cns_CountryPopulation> CreateCountryPopulationAsync(cns_CountryPopulation entity);
        Task<bool> DeleteCountryPopulationByIdAsync(int id);

        // cns_CountryStates
        Task<List<cns_CountryStates>> GetCountryStatesAsync();
        Task<cns_CountryStates?> GetCountryStatesByIdAsync(int id);
        Task<cns_CountryStates> CreateCountryStatesAsync(cns_CountryStates entity);
        Task<bool> DeleteCountryStatesByIdAsync(int id);

        // Bulk operations
        Task ClearAllAsync();
    }
}
