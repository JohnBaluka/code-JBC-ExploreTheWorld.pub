using Microsoft.JSInterop;

namespace ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Services;

public class ClientCommandHandler
{
    [JSInvokable]
    public Task<string> SayHelloWASM(string name)
    {
        Console.WriteLine($"Invoking SayHelloWASM {name}");
        return Task.FromResult($"Hello WASM, {name} from the InteractiveWebAssembly ClientCommandHandler!");
    }

    [JSInvokable]
    public Task<string> SayHelloHome(string name)
    {
        Console.WriteLine($"ClientCommandHandler.SayHelloHome delegating to Home.SayHelloHome");
        return Pages.Home.SayHelloHome(name);
    }
}
