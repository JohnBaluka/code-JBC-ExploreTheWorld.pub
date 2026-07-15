using System;
using System.IO;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl
{
    public static class ExploreTheWorldSqliteDb
    {
        public static readonly string DefaultConnectionString =
            $"Data Source={Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "JBC", "ExploreTheWorld", "etw.db")}";

        public static ExploreTheWorldDbContextFactory CreateFactory(string connectionString)
        {
            var options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
                .UseSqlite(connectionString)
                .Options;
            return new ExploreTheWorldDbContextFactory(options);
        }

        public static CountriesNowSpaceApiManager__Repo CreateRepo(ExploreTheWorldDbContextFactory factory)
            => new CountriesNowSpaceApiManager__Repo(factory);
    }
}
