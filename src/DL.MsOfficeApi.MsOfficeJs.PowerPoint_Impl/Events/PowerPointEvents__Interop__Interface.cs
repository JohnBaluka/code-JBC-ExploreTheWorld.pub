namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl
{
    // Wraps wwwroot/js/events.js (Office.js PowerPoint event registration).
    // dotNetRef must expose [JSInvokable] OnEventLogged(string eventName, string timestamp).
    public interface PowerPointEvents__Interop__Interface
    {
        Task StartWatchingAsync(object dotNetRef, string[] eventKeys);
        Task StopWatchingAsync();
    }
}
