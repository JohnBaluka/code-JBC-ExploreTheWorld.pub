using Microsoft.EntityFrameworkCore;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;

namespace JBC.ExploreTheWorld.IntegrationTests.Fixtures;

// Adapts DbContextOptions into IDbContextFactory<T> for EF Core 3.x, which has CreateDbContext()
// but not the async factory or DI-registered factory from EF Core 5+.
public sealed class TestDbContextFactory : IDbContextFactory<ExploreTheWorldDbContext>
{
    private readonly DbContextOptions<ExploreTheWorldDbContext> _options;

    public TestDbContextFactory(DbContextOptions<ExploreTheWorldDbContext> options)
    {
        _options = options;
    }

    public ExploreTheWorldDbContext CreateDbContext() => new ExploreTheWorldDbContext(_options);
}
