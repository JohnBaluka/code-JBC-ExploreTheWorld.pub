using Microsoft.EntityFrameworkCore;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData
{
    // EF Core 3.1-compatible IDbContextFactory implementation for the _netF project.
    // Replaces the AddDbContextFactory DI registration used in the net10.0 build.
    // Provider-specific options are built by each DL.CountriesNowSpaceData.{Provider}._netF project.
    public sealed class ExploreTheWorldDbContextFactory : IDbContextFactory<ExploreTheWorldDbContext>
    {
        private readonly DbContextOptions<ExploreTheWorldDbContext> _options;

        public ExploreTheWorldDbContextFactory(DbContextOptions<ExploreTheWorldDbContext> options)
        {
            _options = options;
        }

        public ExploreTheWorldDbContext CreateDbContext() =>
            new ExploreTheWorldDbContext(_options);
    }
}
