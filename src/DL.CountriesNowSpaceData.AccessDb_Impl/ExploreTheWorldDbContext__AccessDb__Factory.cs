using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;

internal sealed class ExploreTheWorldDbContext__AccessDb__Factory(DbContextOptions<ExploreTheWorldDbContext> options)
    : IDbContextFactory<ExploreTheWorldDbContext>
{
    public ExploreTheWorldDbContext CreateDbContext() => new(options);
}
