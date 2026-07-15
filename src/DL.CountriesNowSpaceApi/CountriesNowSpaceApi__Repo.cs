using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    public class CountriesNowSpaceApi__Repo : CountriesNowSpaceApi_Interface
    {
        private const string BaseUrl = "https://countriesnow.space/api/v0.1/";

        private readonly HttpClient _client;

        public CountriesNowSpaceApi__Repo() : this(new HttpClient()) { }

        public CountriesNowSpaceApi__Repo(HttpClient client)
        {
            _client = client;
            if (_client.BaseAddress == null)
                _client.BaseAddress = new Uri(BaseUrl);
        }

        public async Task<List<CountryBasic_Row>> GetAllCountriesAsync()
        {
            var response = await _client.GetFromJsonAsync<CountriesNowResponse_Row<List<CountryBasic_Row>>>(
                "countries");
            return response?.Data ?? new List<CountryBasic_Row>();
        }

        public async Task<CountryCapital_Row?> GetCountryCapitalAsync(string country)
        {
            var response = await _client.PostAsJsonAsync("countries/capital", new { country });
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<CountriesNowResponse_Row<CountryCapital_Row>>();
            return result?.Data;
        }

        public async Task<CountryPopulation_Row?> GetCountryPopulationAsync(string country)
        {
            var response = await _client.PostAsJsonAsync("countries/population", new { country });
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<CountriesNowResponse_Row<CountryPopulation_Row>>();
            return result?.Data;
        }

        public async Task<CountryStates_Row?> GetCountryStatesAsync(string country)
        {
            var response = await _client.PostAsJsonAsync("countries/states", new { country });
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<CountriesNowResponse_Row<CountryStates_Row>>();
            return result?.Data;
        }

        public async Task<List<CountryFlag_Row>> GetCountryFlagsAsync()
        {
            var response = await _client.GetFromJsonAsync<CountriesNowResponse_Row<List<CountryFlag_Row>>>(
                "countries/flag/images");
            return response?.Data ?? new List<CountryFlag_Row>();
        }
    }
}
