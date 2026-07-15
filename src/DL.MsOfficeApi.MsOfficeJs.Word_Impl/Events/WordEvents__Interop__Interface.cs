namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl
{
    // Wraps wwwroot/js/events.js (Office.js Word document event registration).
    // dotNetRef must expose [JSInvokable] OnEventLogged(string eventName, string timestamp).
    public interface WordEvents__Interop__Interface
    {
        Task StartWatchingAsync(object dotNetRef, string[] eventKeys);
        Task StopWatchingAsync();
    }
}
