using Microsoft.JSInterop;
using System.Text.Json;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData.IndexedDb_Impl
{
    // Browser IndexedDB key-value service backed by the indexedDb.js ES module.
    // Follows the lazy-load-with-SemaphoreSlim-lock pattern used across DL.MsJSInterop repos.
    public class IndexedDb_Repo : IndexedDb_Repo__Interface, IAsyncDisposable
    {
        private readonly IJSRuntime _js;
        private IJSObjectReference? _module;
        private readonly SemaphoreSlim _moduleLock = new(1, 1);

        private static readonly string[] ModulePathTemplates =
        [
            "./_content/JBC.ExploreTheWorld.DL.CountriesNowSpaceData.IndexedDb_Impl/js/indexedDb.js",
            "/js/indexedDb.js",
        ];

        public IndexedDb_Repo(IJSRuntime js)
        {
            _js = js;
        }

        private async Task<IJSObjectReference> GetModuleAsync()
        {
            await _moduleLock.WaitAsync();
            try
            {
                if (_module is not null) return _module;

                foreach (var path in ModulePathTemplates)
                {
                    try
                    {
                        _module = await _js.InvokeAsync<IJSObjectReference>("import", path);
                        return _module;
                    }
                    catch (JSException) { }
                }

                throw new InvalidOperationException(
                    "Failed to load the IndexedDb JS module from any known path.");
            }
            finally
            {
                _moduleLock.Release();
            }
        }

        public async Task<T?> GetItemAsync<T>(string key, CancellationToken cancellationToken = default)
        {
            var module = await GetModuleAsync();
            var json = await module.InvokeAsync<string?>("getItem", cancellationToken, key);
            if (json is null) return default;
            return JsonSerializer.Deserialize<T>(json);
        }

        public async Task SetItemAsync<T>(string key, T value, CancellationToken cancellationToken = default)
        {
            var module = await GetModuleAsync();
            var json = JsonSerializer.Serialize(value);
            await module.InvokeVoidAsync("setItem", cancellationToken, key, json);
        }

        public async Task RemoveItemAsync(string key, CancellationToken cancellationToken = default)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("removeItem", cancellationToken, key);
        }

        public async ValueTask DisposeAsync()
        {
            if (_module is not null)
            {
                try { await _module.DisposeAsync(); } catch { }
            }
            _moduleLock.Dispose();
        }
    }
}
