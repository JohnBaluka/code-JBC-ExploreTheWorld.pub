using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.IntegrationTests.Fixtures;

public class TestDbContextFactory : IDbContextFactory<ExploreTheWorldDbContext>
{
    private readonly DbContextOptions<ExploreTheWorldDbContext> _options;

    public TestDbContextFactory(DbContextOptions<ExploreTheWorldDbContext> options)
        => _options = options;

    public ExploreTheWorldDbContext CreateDbContext()
        => new ExploreTheWorldDbContext(_options);

    public Task<ExploreTheWorldDbContext> CreateDbContextAsync(CancellationToken cancellationToken = default)
        => Task.FromResult(CreateDbContext());
}
