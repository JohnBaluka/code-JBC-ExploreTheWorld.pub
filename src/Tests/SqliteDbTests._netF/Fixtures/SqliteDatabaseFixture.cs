using Microsoft.EntityFrameworkCore;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.SqliteDbTests.Fixtures;

/// <summary>
/// Creates a real SQLite database file in the temp folder (EF Core 3.x provider),
/// seeds it, and deletes the file on disposal. One instance per test class.
/// </summary>
public class SqliteDatabaseFixture : IDisposable
{
    private readonly string _databasePath;
    private readonly DbContextOptions<ExploreTheWorldDbContext> _options;

    public SqliteDatabaseFixture()
    {
        _databasePath = Path.Combine(Path.GetTempPath(), $"ETW_SqliteDbTests_{Guid.NewGuid():N}.db");

        _options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
            .UseSqlite($"Data Source={_databasePath}")
            .Options;

        using var context = new ExploreTheWorldDbContext(_options);
        context.Database.EnsureCreated();
        SeedData(context);
    }

    public ExploreTheWorldDbContextFactory CreateFactory()
        => new ExploreTheWorldDbContextFactory(_options);

    private static void SeedData(ExploreTheWorldDbContext context)
    {
        context.cns_Countries.AddRange(
            new cns_Country { Iso2 = "AU", Country = "Australia", Iso3 = "AUS" },
            new cns_Country { Iso2 = "BR", Country = "Brazil",    Iso3 = "BRA" },
            new cns_Country { Iso2 = "CA", Country = "Canada",    Iso3 = "CAN" }
        );

        context.SaveChanges();
    }

    public void Dispose()
    {
        // Microsoft.Data.Sqlite 3.x has no connection pooling, so once the contexts are
        // disposed the file can be deleted directly.
        try { if (File.Exists(_databasePath)) File.Delete(_databasePath); } catch { }
    }
}
