using Microsoft.EntityFrameworkCore;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;

namespace JBC.ExploreTheWorld.IntegrationTests.Fixtures;

public class DatabaseFixture : IDisposable
{
    private readonly DbContextOptions<ExploreTheWorldDbContext> _options;

    public DatabaseFixture()
    {
        var dbName = $"ExploreTheWorld_Test_{Guid.NewGuid():N}";

        _options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
            .UseInMemoryDatabase(dbName)
            .Options;

        using var context = CreateFreshContext();
        SeedData(context);
    }

    public ExploreTheWorldDbContext CreateFreshContext()
        => new ExploreTheWorldDbContext(_options);

    public IDbContextFactory<ExploreTheWorldDbContext> CreateFactory()
        => new TestDbContextFactory(_options);

    private static void SeedData(ExploreTheWorldDbContext context)
    {
        context.cns_Countries.AddRange(
            new cns_Country { Iso2 = "AU", Country = "Australia", Iso3 = "AUS" },
            new cns_Country { Iso2 = "BR", Country = "Brazil",    Iso3 = "BRA" },
            new cns_Country { Iso2 = "CA", Country = "Canada",    Iso3 = "CAN" }
        );

        context.SaveChanges();
    }

    public void Dispose() { }
}
