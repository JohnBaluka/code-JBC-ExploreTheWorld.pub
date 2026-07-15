namespace JBC.ExploreTheWorld.DL
{
    // Canonical provider-name keys shared by the DB-provider _Impl registrations, the BL
    // DbProviderSwitcher__Service, and every host's AvailableProviders list / DbProvider_AppService seed.
    // Centralized here (core DL) so the switch keys are defined once instead of duplicated as
    // magic strings across projects.
    public static class DbProviderNames
    {
        // Server-side EF Core providers
        public const string SqliteDb = "SqliteDb";
        public const string SqlServerDb = "SqlServerDb";
        public const string InMemoryDb = "InMemoryDb";
        public const string AccessDb = "AccessDb";

        // Browser storage providers (Blazor WebAssembly)
        public const string LocalStorageDb = "LocalStorageDb";
        public const string SessionStorageDb = "SessionStorageDb";
        public const string IndexedDb = "IndexedDb";
    }
}
