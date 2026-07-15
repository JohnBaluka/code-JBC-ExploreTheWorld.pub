using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;

public sealed class CountriesNowSpaceApiManager__InMemoryDb__Repo(IDbContextFactory<ExploreTheWorldDbContext> factory)
    : CountriesNowSpaceApiManager__Repo(factory)
{ }
