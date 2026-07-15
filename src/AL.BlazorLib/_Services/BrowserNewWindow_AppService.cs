using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.AL.BlazorLib
{
    public class BrowserNewWindow_AppService : NewWindow_AppService__Interface
    {
        private readonly IJSRuntime _js;

        public BrowserNewWindow_AppService(IJSRuntime js) { _js = js; }

        public void OpenNewWindow()
        {
            _ = _js.InvokeVoidAsync("open", "/", "_blank");
        }
    }
}
