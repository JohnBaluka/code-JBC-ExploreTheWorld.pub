using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;

internal sealed class ExploreTheWorldDbContext__SqliteDb__Factory(DbContextOptions<ExploreTheWorldDbContext> options)
    : IDbContextFactory<ExploreTheWorldDbContext>
{
    public ExploreTheWorldDbContext CreateDbContext() => new(options);
}
