using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl;

public sealed class CountriesNowSpaceApiManager__SqlServerDb__Repo(IDbContextFactory<ExploreTheWorldDbContext> factory)
    : CountriesNowSpaceApiManager__Repo(factory)
{ }
