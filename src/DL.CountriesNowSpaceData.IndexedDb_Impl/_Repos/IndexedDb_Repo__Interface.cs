namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.IndexedDb_Impl
{
    public interface IndexedDb_Repo__Interface
    {
        Task<T?> GetItemAsync<T>(string key, CancellationToken cancellationToken = default);
        Task SetItemAsync<T>(string key, T value, CancellationToken cancellationToken = default);
        Task RemoveItemAsync(string key, CancellationToken cancellationToken = default);
    }
}
