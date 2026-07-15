using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData
{
    // EF Core 3.1-compatible helper to ensure the database schema exists at startup.
    // Called directly (no DI IServiceProvider) in the _netF non-DI host pattern.
    public static class ServiceCollectionExtensions
    {
        public static void EnsureExploreTheWorldDbCreated(
            ExploreTheWorldDbContextFactory factory)
        {
            using var context = factory.CreateDbContext();
            context.Database.EnsureCreated();
        }
    }
}
