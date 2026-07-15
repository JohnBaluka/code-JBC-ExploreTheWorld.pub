using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;

internal sealed class ExploreTheWorldDbContext__InMemoryDb__Factory(DbContextOptions<ExploreTheWorldDbContext> options)
    : IDbContextFactory<ExploreTheWorldDbContext>
{
    public ExploreTheWorldDbContext CreateDbContext() => new(options);
}
