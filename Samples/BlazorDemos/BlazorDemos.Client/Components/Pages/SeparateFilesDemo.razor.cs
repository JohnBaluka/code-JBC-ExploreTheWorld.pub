using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorDemos.Client.Components.Pages;

public partial class SeparateFilesDemo : IAsyncDisposable
{
    [Inject]
    private IJSRuntime JS { get; set; } = default!;

    private string Name { get; set; } = "";
    private string Message { get; set; } = "Enter a name and click the button.";

    private IJSObjectReference? _module;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            _module = await JS.InvokeAsync<IJSObjectReference>(
                "import",
                "./Components/Pages/SeparateFilesDemo.razor.js");
        }
    }

    private async Task ShowGreetingAsync()
    {
        if (string.IsNullOrWhiteSpace(Name))
        {
            Message = "Please enter a name.";
            return;
        }

        Message = $"Hello, {Name}!";

        if (_module is not null)
        {
            await _module.InvokeVoidAsync("showAlert", Message);
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_module is not null)
        {
            await _module.DisposeAsync();
        }
    }
}
