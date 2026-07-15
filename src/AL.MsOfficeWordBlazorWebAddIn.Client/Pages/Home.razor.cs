using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

using ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client.Model;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class Home : ComponentBase
{
    // home.js now lives in DL.MsOfficeApi.MsOfficeJs.Word_Impl and is served as a static web asset.
    private const string HomeModulePath =
        "_content/JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl/js/home.js";

    private HostInformation _hostInformation = new();

    [Inject]
    private IJSRuntime JSRuntime { get; set; } = default!;

    [Inject]
    private NavigationManager Navigation { get; set; } = default!;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            _hostInformation = await JSRuntime.InvokeAsync<HostInformation>("Office.onReady");

            if (_hostInformation.IsInitialized)
            {
                StateHasChanged();
            }
        }
    }

    [JSImport("sayHelloFromJs", "Home")]
    internal static partial string SayHelloFromJsFunction(string name);

    [JSInvokable]
    public static async Task<string> SayHelloHome(string name)
    {
        Console.WriteLine($"Invoking static SayHelloHome {name}");
        await JSHost.ImportAsync("Home", HomeModulePath);
        var result = SayHelloFromJsFunction(name);
        var renderMode = OperatingSystem.IsBrowser() ? "InteractiveWebAssembly" : "Unknown";
        return $"{result} from the {renderMode} Home Page!";
    }
}
