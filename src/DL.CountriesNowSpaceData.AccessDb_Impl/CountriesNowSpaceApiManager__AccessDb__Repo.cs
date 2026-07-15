using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;

public sealed class CountriesNowSpaceApiManager__AccessDb__Repo(IDbContextFactory<ExploreTheWorldDbContext> factory)
    : CountriesNowSpaceApiManager__Repo(factory)
{ }
