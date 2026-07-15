using System;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl
{
    public static class ExploreTheWorldSqlServerDb
    {
        public static readonly string DefaultConnectionString =
            "Server=.;Database=JBC_ExploreTheWorld;Trusted_Connection=True;";

        public static ExploreTheWorldDbContextFactory CreateFactory(string connectionString)
        {
            var options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
                .UseSqlServer(connectionString)
                .Options;
            return new ExploreTheWorldDbContextFactory(options);
        }

        public static CountriesNowSpaceApiManager__Repo CreateRepo(ExploreTheWorldDbContextFactory factory)
            => new CountriesNowSpaceApiManager__Repo(factory);
    }
}
