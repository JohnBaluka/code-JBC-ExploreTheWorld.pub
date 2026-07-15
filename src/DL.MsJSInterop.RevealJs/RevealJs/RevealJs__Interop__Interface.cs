using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs
{
    public interface RevealJs__Interop__Interface
    {
        bool IsInitialized { get; }

        Task<bool> InitializeAsync(object container, object dotNetRef, string transition, string theme, int logLevel);
        Task DestroyAsync();

        Task NavigateNextAsync();
        Task NavigatePrevAsync();
        Task NavigateRightAsync();
        Task NavigateLeftAsync();
        Task NavigateToSlideAsync(int h, int v);

        Task SetTransitionAsync(string transition);
        Task SetThemeAsync(string theme);
        Task SetSlideNumberAsync(bool show);

        Task StartAutoPlayAsync(int intervalMs);
        Task StopAutoPlayAsync();

        Task ToggleOverviewAsync();
        Task TogglePauseAsync();

        Task SetNotesPositionAsync(string position);
        Task ToggleSearchAsync();

        Task RequestFullscreenAsync(object container);
        Task<bool> IsFullscreenAsync();

        Task LayoutAsync();
    }
}
