using ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Components;

public partial class ServerBridge : ComponentBase, IDisposable
{
    [Inject]
    private IJSRuntime JSRuntime { get; set; } = default!;

    [Inject]
    private ServerCommandHandler Handler { get; set; } = default!;

    private DotNetObjectReference<ServerCommandHandler>? _dotNetRef;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            Console.WriteLine("ServerBridge: Registering ServerCommandHandler with JavaScript");
            _dotNetRef = DotNetObjectReference.Create(Handler);
            try
            {
                await JSRuntime.InvokeVoidAsync("window.signalDotNetReady", "server", _dotNetRef);
            }
            catch (JSException ex)
            {
                Console.WriteLine($"ServerBridge: Error signaling .NET ready: {ex.Message}");
                throw;
            }
        }
    }

    public void Dispose()
    {
        _dotNetRef?.Dispose();
    }
}
