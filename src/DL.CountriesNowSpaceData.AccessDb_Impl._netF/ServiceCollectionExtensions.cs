using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl
{
    public static class ExploreTheWorldAccessDb
    {
        public static readonly string DefaultDbPath =
            @"C:\Dev\github\JohnBaluka\code-JBC-ExploreTheWorld\VBA\Access\ExploreTheWorld.accdb";

        public static ExploreTheWorldDbContextFactory CreateFactory(string dbPath)
        {
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";
            var options = new DbContextOptionsBuilder<ExploreTheWorldDbContext>()
                .UseJet(connectionString)
                .Options;
            return new ExploreTheWorldDbContextFactory(options);
        }

        public static CountriesNowSpaceApiManager__Repo CreateRepo(ExploreTheWorldDbContextFactory factory)
            => new CountriesNowSpaceApiManager__Repo(factory);
    }
}
