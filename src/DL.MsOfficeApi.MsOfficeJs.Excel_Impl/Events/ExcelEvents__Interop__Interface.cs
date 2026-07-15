namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl
{
    // Wraps wwwroot/js/events.js (Office.js Excel worksheet event registration).
    // dotNetRef must expose [JSInvokable] OnEventLogged(string eventName, string timestamp).
    public interface ExcelEvents__Interop__Interface
    {
        Task StartWatchingAsync(object dotNetRef, string[] eventKeys);
        Task StopWatchingAsync();
    }
}
