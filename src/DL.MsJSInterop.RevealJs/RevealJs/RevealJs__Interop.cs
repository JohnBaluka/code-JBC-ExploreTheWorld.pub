using Microsoft.JSInterop;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs
{
    public class RevealJs__Interop : JsModuleInterop__Base, RevealJs__Interop__Interface
    {
        public RevealJs__Interop(IJSRuntime jsRuntime)
            : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs/js/reveal-interop.js")
        {
        }

        public async Task<bool> InitializeAsync(object container, object dotNetRef, string transition, string theme, int logLevel)
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<bool>("initialize", container, dotNetRef, transition, theme, logLevel);
        }

        public async Task DestroyAsync()
        {
            if (!IsInitialized) return;
            var module = await GetModuleAsync();
            try { await module.InvokeVoidAsync("destroy"); } catch { }
        }

        public async Task NavigateNextAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("navigateNext");
        }

        public async Task NavigatePrevAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("navigatePrev");
        }

        public async Task NavigateRightAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("navigateRight");
        }

        public async Task NavigateLeftAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("navigateLeft");
        }

        public async Task NavigateToSlideAsync(int h, int v)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("navigateToSlide", h, v);
        }

        public async Task SetTransitionAsync(string transition)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("setTransition", transition);
        }

        public async Task SetThemeAsync(string theme)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("setTheme", theme);
        }

        public async Task SetSlideNumberAsync(bool show)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("setSlideNumber", show);
        }

        public async Task StartAutoPlayAsync(int intervalMs)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("startAutoPlay", intervalMs);
        }

        public async Task StopAutoPlayAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("stopAutoPlay");
        }

        public async Task ToggleOverviewAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("toggleOverview");
        }

        public async Task TogglePauseAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("togglePause");
        }

        public async Task SetNotesPositionAsync(string position)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("setNotesPosition", position);
        }

        public async Task ToggleSearchAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("toggleSearch");
        }

        public async Task RequestFullscreenAsync(object container)
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("requestFullscreen", container);
        }

        public async Task<bool> IsFullscreenAsync()
        {
            var module = await GetModuleAsync();
            return await module.InvokeAsync<bool>("isFullscreen");
        }

        public async Task LayoutAsync()
        {
            var module = await GetModuleAsync();
            await module.InvokeVoidAsync("layout");
        }
    }
}
