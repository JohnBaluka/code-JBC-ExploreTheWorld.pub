using ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Components;

public partial class WasmBridge : ComponentBase, IDisposable
{
    [Inject]
    private IJSRuntime JSRuntime { get; set; } = default!;

    [Inject]
    private ClientCommandHandler Handler { get; set; } = default!;

    private DotNetObjectReference<ClientCommandHandler>? _dotNetRef;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            Console.WriteLine("WasmBridge: Registering ClientCommandHandler with JavaScript");
            _dotNetRef = DotNetObjectReference.Create(Handler);
            try
            {
                await JSRuntime.InvokeVoidAsync("window.signalDotNetReady", "wasm", _dotNetRef);
            }
            catch (JSException ex)
            {
                Console.WriteLine($"WasmBridge: Error signaling .NET ready: {ex.Message}");
                throw;
            }
        }
    }

    public void Dispose()
    {
        _dotNetRef?.Dispose();
    }
}
