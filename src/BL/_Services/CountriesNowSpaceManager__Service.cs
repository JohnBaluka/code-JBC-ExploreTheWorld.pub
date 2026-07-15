using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.BL
{
    /// <summary>
    /// Returns CountriesNow.space data with a DB-backed cache-first strategy: read from the
    /// database when data is cached; otherwise fetch from the API, persist, and return. The
    /// data dependencies (<see cref="CountriesNowSpaceApi_Interface"/> and
    /// <see cref="CountriesNowSpaceApiManager__Repo__Interface"/>) are DL interfaces, so the
    /// service is exercised in unit tests by mocking them directly — it needs no interface of its own.
    /// </summary>
    public class CountriesNowSpaceManager__Service
    {
        private readonly CountriesNowSpaceApi_Interface _apiService;
        private readonly CountriesNowSpaceApiManager__Repo__Interface _dbManager;

        public CountriesNowSpaceManager__Service(
            CountriesNowSpaceApi_Interface apiService,
            CountriesNowSpaceApiManager__Repo__Interface dbManager)
        {
            _apiService = apiService;
            _dbManager  = dbManager;
        }

        // ── GetAllCountriesAsync ─────────────────────────────────────────────────────

        public async Task<DataResult_Row<CountryBasic_Row>> GetAllCountriesAsync()
        {
            var dbCountries = await _dbManager.GetCountriesAsync();

            if (dbCountries.Count > 0)
            {
                return new DataResult_Row<CountryBasic_Row>
                {
                    Source = DataSource_Enum.Database,
                    Data   = dbCountries
                        .Select(c => new CountryBasic_Row
                        {
                            Country = c.Country,
                            Iso2    = c.Iso2,
                            Iso3    = c.Iso3
                        })
                        .ToList()
                };
            }

            var apiCountries = await _apiService.GetAllCountriesAsync();

            foreach (var c in apiCountries)
            {
                if (string.IsNullOrEmpty(c.Iso2)) continue;

                await _dbManager.CreateCountryAsync(new cns_Country
                {
                    Iso2    = c.Iso2,
                    Country = c.Country,
                    Iso3    = c.Iso3
                });
            }

            return new DataResult_Row<CountryBasic_Row>
            {
                Source = DataSource_Enum.Api,
                Data   = apiCountries
            };
        }

        // ── GetCountryStatesAsync ────────────────────────────────────────────────────

        public async Task<DataResult_Row<CountryState_Row>> GetCountryStatesAsync(string country)
        {
            var allDbStates = await _dbManager.GetCountryStatesAsync();
            var dbEntry = allDbStates.FirstOrDefault(s =>
                string.Equals(s.Name, country, StringComparison.OrdinalIgnoreCase));

            if (dbEntry != null)
            {
                return new DataResult_Row<CountryState_Row>
                {
                    Source = DataSource_Enum.Database,
                    Data   = dbEntry.States
                        .Select(s => new CountryState_Row
                        {
                            Name      = s.Name,
                            StateCode = s.StateCode
                        })
                        .ToList()
                };
            }

            var apiStates = await _apiService.GetCountryStatesAsync(country);
            var stateList = apiStates?.States ?? new List<CountryState_Row>();

            if (apiStates != null)
            {
                var entity = new cns_CountryStates
                {
                    Name   = apiStates.Name,
                    Iso3   = apiStates.Iso3,
                    States = stateList
                        .Select(s => new cns_CountryState
                        {
                            Name      = s.Name,
                            StateCode = s.StateCode
                        })
                        .ToList()
                };
                await _dbManager.CreateCountryStatesAsync(entity);
            }

            return new DataResult_Row<CountryState_Row>
            {
                Source = DataSource_Enum.Api,
                Data   = stateList
            };
        }

        // ── GetCountryFlagsAsync ─────────────────────────────────────────────────────

        public async Task<DataResult_Row<CountryFlag_Row>> GetCountryFlagsAsync()
        {
            var dbFlags = await _dbManager.GetCountryFlagsAsync();

            if (dbFlags.Count > 0)
            {
                return new DataResult_Row<CountryFlag_Row>
                {
                    Source = DataSource_Enum.Database,
                    Data   = dbFlags
                        .Select(f => new CountryFlag_Row
                        {
                            Iso2     = f.Iso2,
                            Name     = f.Name,
                            Flag     = f.Flag,
                            DialCode = f.DialCode
                        })
                        .ToList()
                };
            }

            var apiFlags = await _apiService.GetCountryFlagsAsync();

            var seenIso2 = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in apiFlags)
            {
                if (string.IsNullOrEmpty(f.Iso2) || !seenIso2.Add(f.Iso2)) continue;

                await _dbManager.CreateCountryFlagAsync(new cns_CountryFlag
                {
                    Iso2     = f.Iso2,
                    Name     = f.Name,
                    Flag     = f.Flag,
                    DialCode = f.DialCode
                });
            }

            return new DataResult_Row<CountryFlag_Row>
            {
                Source = DataSource_Enum.Api,
                Data   = apiFlags
            };
        }

        // ── ClearAllDataAsync ────────────────────────────────────────────────────────

        public Task ClearAllDataAsync() => _dbManager.ClearAllAsync();
    }
}
