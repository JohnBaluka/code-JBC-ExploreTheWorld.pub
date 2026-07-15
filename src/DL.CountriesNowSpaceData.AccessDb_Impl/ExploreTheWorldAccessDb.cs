namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl
{
    public static class ExploreTheWorldAccessDb
    {
        // Separate EF Core data file; never the VBA application database.
        public static readonly string DefaultDbPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "JBC.ExploreTheWorld", "etw.accdb");
    }
}
