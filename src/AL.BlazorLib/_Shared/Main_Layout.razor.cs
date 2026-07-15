using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;

namespace JBC.ExploreTheWorld.AL.BlazorLib
{
    public partial class Main_Layout : IAsyncDisposable
    {
        [Inject] private Layout__Interop__Interface LayoutInterop { get; set; } = default!;
        [Inject] private Layout_AppService Layout_AppService { get; set; } = default!;
        [Inject] private DbProvider_AppService DbProvider_AppService { get; set; } = default!;
        [Inject] private NewWindow_AppService__Interface NewWindow_AppService { get; set; } = default!;

        private bool _sidebarExpanded = true;
        private string _currentBreakpoint = "—";
        private string _selectedDbProvider = string.Empty;
        private DotNetObjectReference<Main_Layout>? _selfRef;
        private IJSObjectReference? _resizeSubscription;

        // RendererInfo is unavailable in Radzen Blazor Studio's design-time preview renderer
        // (its RenderHandle throws "No renderer has been initialized"); fall back to a neutral
        // label there. In the real hosts RendererInfo is always available, so this returns the
        // normal Server/WebAssembly/Static SSR value.
        private string RenderModeDisplay
        {
            get
            {
                try
                {
                    return RendererInfo.IsInteractive ? RendererInfo.Name : "Static SSR";
                }
                catch (InvalidOperationException)
                {
                    return "Preview";
                }
            }
        }

        private IReadOnlyList<string> DropdownProviders =>
            DbProvider_AppService.AvailableProviders.Count > 0
                ? DbProvider_AppService.AvailableProviders
                : (IReadOnlyList<string>)[DbProvider_AppService.ProviderName];

        protected override void OnInitialized()
        {
            _selectedDbProvider = DbProvider_AppService.ProviderName;
            DbProvider_AppService.OnProviderChanged += OnProviderChangedExternally;
        }

        private void OnDbProviderChanged(string value)
        {
            _selectedDbProvider = value;
            DbProvider_AppService.SetProvider(value);
        }

        private void OnProviderChangedExternally()
        {
            _selectedDbProvider = DbProvider_AppService.ProviderName;
            InvokeAsync(StateHasChanged);
        }

        private BadgeStyle GetBreakpointBadgeStyle() => _currentBreakpoint switch
        {
            "XS"  => BadgeStyle.Danger,
            "SM"  => BadgeStyle.Warning,
            "MD"  => BadgeStyle.Info,
            "LG"  => BadgeStyle.Success,
            "XL"  => BadgeStyle.Primary,
            _     => BadgeStyle.Secondary
        };

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                try
                {
                    var width = await LayoutInterop.GetWindowWidthAsync();
                    UpdateBreakpoint(width);
                    _selfRef = DotNetObjectReference.Create(this);
                    _resizeSubscription = await LayoutInterop.WatchWindowWidthAsync(_selfRef);
                    StateHasChanged();
                }
                catch (JSException) { }
            }
        }

        private void OpenNewWindow() => NewWindow_AppService.OpenNewWindow();

        [JSInvokable]
        public void OnWindowWidthChanged(int width)
        {
            UpdateBreakpoint(width);
            StateHasChanged();
        }

        private void UpdateBreakpoint(int width)
        {
            _currentBreakpoint = width switch
            {
                < 576  => "XS",
                < 768  => "SM",
                < 992  => "MD",
                < 1200 => "LG",
                < 1400 => "XL",
                _      => "XXL"
            };
        }

        public async ValueTask DisposeAsync()
        {
            DbProvider_AppService.OnProviderChanged -= OnProviderChangedExternally;

            if (_resizeSubscription is not null)
            {
                // Both the "dispose" call and DisposeAsync() issue JS interop. When the circuit
                // has already gone (F5 refresh, Server↔WASM switch, tab close) these throw
                // JSDisconnectedException; there is nothing left to clean up on the client, so
                // swallow disconnect/interop failures. Other exceptions still propagate.
                try
                {
                    await _resizeSubscription.InvokeVoidAsync("dispose");
                    await _resizeSubscription.DisposeAsync();
                }
                catch (JSDisconnectedException) { }
                catch (JSException) { }
                catch (OperationCanceledException) { }
            }
            _selfRef?.Dispose();
        }
    }
}
