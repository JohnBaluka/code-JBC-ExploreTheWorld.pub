using Microsoft.EntityFrameworkCore;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData
{
    public class CountriesNowSpaceApiManager__Repo : CountriesNowSpaceApiManager__Repo__Interface
    {
        private readonly IDbContextFactory<ExploreTheWorldDbContext> _contextFactory;

        public CountriesNowSpaceApiManager__Repo(IDbContextFactory<ExploreTheWorldDbContext> contextFactory)
        {
            _contextFactory = contextFactory;
        }

        // ================================================================
        // cns_Country
        // ================================================================

        public async Task<List<cns_Country>> GetCountriesAsync()
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_Countries
                .Include(e => e.Cities)
                .Include(e => e.Capital)
                .Include(e => e.Flag)
                .ToListAsync();
        }

        public async Task<cns_Country?> GetCountryByIso2Async(string iso2)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_Countries
                .Include(e => e.Cities)
                .Include(e => e.Capital)
                .Include(e => e.Flag)
                .FirstOrDefaultAsync(e => e.Iso2 == iso2);
        }

        public async Task<cns_Country> CreateCountryAsync(cns_Country entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_Countries.Add(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<cns_Country> UpdateCountryAsync(cns_Country entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_Countries.Update(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<bool> DeleteCountryByIso2Async(string iso2)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            var entity = await context.cns_Countries.FindAsync(iso2);
            if (entity == null) return false;
            context.cns_Countries.Remove(entity);
            await context.SaveChangesAsync();
            return true;
        }

        // ================================================================
        // cns_City
        // ================================================================

        public async Task<List<cns_City>> GetCitiesByIso2Async(string iso2)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_Cities
                .Where(e => e.Iso2 == iso2)
                .ToListAsync();
        }

        public async Task<cns_City> CreateCityAsync(cns_City entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_Cities.Add(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<bool> DeleteCitiesByIso2Async(string iso2)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            var entities = await context.cns_Cities.Where(e => e.Iso2 == iso2).ToListAsync();
            if (entities.Count == 0) return false;
            context.cns_Cities.RemoveRange(entities);
            await context.SaveChangesAsync();
            return true;
        }

        // ================================================================
        // cns_CountryCapital
        // ================================================================

        public async Task<cns_CountryCapital?> GetCountryCapitalByIso2Async(string iso2)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_CountryCapitals.FindAsync(iso2);
        }

        public async Task<cns_CountryCapital> CreateCountryCapitalAsync(cns_CountryCapital entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_CountryCapitals.Add(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<cns_CountryCapital> UpdateCountryCapitalAsync(cns_CountryCapital entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_CountryCapitals.Update(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<bool> DeleteCountryCapitalByIso2Async(string iso2)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            var entity = await context.cns_CountryCapitals.FindAsync(iso2);
            if (entity == null) return false;
            context.cns_CountryCapitals.Remove(entity);
            await context.SaveChangesAsync();
            return true;
        }

        // ================================================================
        // cns_CountryFlag
        // ================================================================

        public async Task<List<cns_CountryFlag>> GetCountryFlagsAsync()
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_CountryFlags.ToListAsync();
        }

        public async Task<cns_CountryFlag?> GetCountryFlagByIso2Async(string iso2)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_CountryFlags.FindAsync(iso2);
        }

        public async Task<cns_CountryFlag> CreateCountryFlagAsync(cns_CountryFlag entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_CountryFlags.Add(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<cns_CountryFlag> UpdateCountryFlagAsync(cns_CountryFlag entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_CountryFlags.Update(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<bool> DeleteCountryFlagByIso2Async(string iso2)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            var entity = await context.cns_CountryFlags.FindAsync(iso2);
            if (entity == null) return false;
            context.cns_CountryFlags.Remove(entity);
            await context.SaveChangesAsync();
            return true;
        }

        // ================================================================
        // cns_CountryPopulation
        // ================================================================

        public async Task<List<cns_CountryPopulation>> GetCountryPopulationsAsync()
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_CountryPopulations
                .Include(e => e.PopulationCounts)
                .ToListAsync();
        }

        public async Task<cns_CountryPopulation?> GetCountryPopulationByIdAsync(int id)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_CountryPopulations
                .Include(e => e.PopulationCounts)
                .FirstOrDefaultAsync(e => e.ID == id);
        }

        public async Task<cns_CountryPopulation> CreateCountryPopulationAsync(cns_CountryPopulation entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_CountryPopulations.Add(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<bool> DeleteCountryPopulationByIdAsync(int id)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            var entity = await context.cns_CountryPopulations.FindAsync(id);
            if (entity == null) return false;
            context.cns_CountryPopulations.Remove(entity);
            await context.SaveChangesAsync();
            return true;
        }

        // ================================================================
        // cns_CountryStates
        // ================================================================

        public async Task<List<cns_CountryStates>> GetCountryStatesAsync()
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_CountryStates
                .Include(e => e.States)
                .ToListAsync();
        }

        public async Task<cns_CountryStates?> GetCountryStatesByIdAsync(int id)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            return await context.cns_CountryStates
                .Include(e => e.States)
                .FirstOrDefaultAsync(e => e.ID == id);
        }

        public async Task<cns_CountryStates> CreateCountryStatesAsync(cns_CountryStates entity)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            context.cns_CountryStates.Add(entity);
            await context.SaveChangesAsync();
            return entity;
        }

        public async Task<bool> DeleteCountryStatesByIdAsync(int id)
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            var entity = await context.cns_CountryStates.FindAsync(id);
            if (entity == null) return false;
            context.cns_CountryStates.Remove(entity);
            await context.SaveChangesAsync();
            return true;
        }

        // ================================================================
        // Bulk operations
        // ================================================================

        public async Task ClearAllAsync()
        {
            using var context = await _contextFactory.CreateDbContextAsync();
            // Delete in dependency order; cascade handles children where FK is required.
            await context.cns_CountryStateItems.ExecuteDeleteAsync();
            await context.cns_CountryStates.ExecuteDeleteAsync();
            await context.cns_PopulationCounts.ExecuteDeleteAsync();
            await context.cns_CountryPopulations.ExecuteDeleteAsync();
            await context.cns_Cities.ExecuteDeleteAsync();
            await context.cns_CountryCapitals.ExecuteDeleteAsync();
            await context.cns_CountryFlags.ExecuteDeleteAsync();
            await context.cns_Countries.ExecuteDeleteAsync();
        }
    }
}