using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;

public sealed class CountriesNowSpaceApiManager__SqliteDb__Repo(IDbContextFactory<ExploreTheWorldDbContext> factory)
    : CountriesNowSpaceApiManager__Repo(factory)
{ }
