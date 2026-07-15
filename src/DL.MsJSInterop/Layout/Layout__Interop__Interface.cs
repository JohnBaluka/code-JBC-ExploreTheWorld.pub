using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.DL.MsJSInterop.Layout
{
    // Generic browser window/layout interop (wraps wwwroot/js/layout.js).
    public interface Layout__Interop__Interface
    {
        Task<int> GetWindowWidthAsync();

        // Returns a subscription handle; invoke "dispose" on it and DisposeAsync()
        // to remove the resize listener. dotNetRef must expose [JSInvokable]
        // OnWindowWidthChanged(int).
        Task<IJSObjectReference> WatchWindowWidthAsync(object dotNetRef);
    }
}
