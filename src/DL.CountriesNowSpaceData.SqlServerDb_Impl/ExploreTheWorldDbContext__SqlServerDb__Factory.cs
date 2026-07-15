using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl;

internal sealed class ExploreTheWorldDbContext__SqlServerDb__Factory(DbContextOptions<ExploreTheWorldDbContext> options)
    : IDbContextFactory<ExploreTheWorldDbContext>
{
    public ExploreTheWorldDbContext CreateDbContext() => new(options);
}
