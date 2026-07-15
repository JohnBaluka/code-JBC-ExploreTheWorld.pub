using System.Collections.Generic;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    public interface CountriesNowSpaceApi_Interface
    {
        Task<List<CountryBasic_Row>> GetAllCountriesAsync();
        Task<CountryCapital_Row?> GetCountryCapitalAsync(string country);
        Task<CountryPopulation_Row?> GetCountryPopulationAsync(string country);
        Task<CountryStates_Row?> GetCountryStatesAsync(string country);
        Task<List<CountryFlag_Row>> GetCountryFlagsAsync();
    }
}
