using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsJSInterop
{
    // Base class for JS-interop services that lazily import an ESM module from a
    // _content/ static-asset path (with cache-busting) and invoke its exports.
    // Derive a typed __Interop service per JS module; register it via DI and inject
    // the __Interop__Interface into Blazor components instead of calling IJSRuntime.
    public abstract class JsModuleInterop__Base : IAsyncDisposable
    {
        private readonly IJSRuntime _jsRuntime;
        private readonly string _modulePath;
        private IJSObjectReference? _module;
        private readonly SemaphoreSlim _moduleLock = new(1, 1);
        private string? _cacheBuster;

        protected JsModuleInterop__Base(IJSRuntime jsRuntime, string modulePath)
        {
            _jsRuntime  = jsRuntime;
            _modulePath = modulePath;
        }

        public bool IsInitialized => _module != null;

        protected async Task<IJSObjectReference> GetModuleAsync()
        {
            if (_module != null) return _module;

            await _moduleLock.WaitAsync();
            try
            {
                if (_module != null) return _module;

                _cacheBuster ??= DateTime.UtcNow.Ticks.ToString();
                _module = await _jsRuntime.InvokeAsync<IJSObjectReference>(
                    "import", $"{_modulePath}?v={_cacheBuster}");
                return _module;
            }
            finally
            {
                _moduleLock.Release();
            }
        }

        public async ValueTask DisposeAsync()
        {
            _moduleLock.Dispose();
            if (_module != null)
            {
                try { await _module.DisposeAsync(); } catch { }
                _module = null;
            }
            GC.SuppressFinalize(this);
        }
    }
}
