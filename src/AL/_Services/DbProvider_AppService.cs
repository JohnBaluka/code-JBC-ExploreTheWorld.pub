using System;
using System.Collections.Generic;

namespace JBC.ExploreTheWorld.AL
{
    /// <summary>
    /// Tracks the active database provider and the set of providers the user may switch between.
    /// Registered as a singleton; the header dropdown binds to <see cref="AvailableProviders"/>.
    /// </summary>
    public class DbProvider_AppService
    {
        public string ProviderName { get; set; } = string.Empty;

        // Non-empty only in browser (WASM) apps where the user can switch providers at runtime.
        // Server-side apps leave this as an empty list; the dropdown is disabled in that case.
        public IReadOnlyList<string> AvailableProviders { get; set; } = new List<string>();

        // Raised when SetProvider changes the active provider; allows layout to re-render.
        public event Action? OnProviderChanged;

        public void SetProvider(string providerName)
        {
            if (providerName == ProviderName) return;
            ProviderName = providerName;
            OnProviderChanged?.Invoke();
        }
    }
}
