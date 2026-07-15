using Microsoft.JSInterop;

namespace ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Services;

public class ServerCommandHandler
{
    [JSInvokable]
    public Task<string> SayHelloServer(string name)
    {
        Console.WriteLine($"Invoking SayHelloServer {name}");
        return Task.FromResult($"Hello Server, {name} from the InteractiveServer ServerCommandHandler!");
    }
}
