using Microsoft.AspNetCore.Components;
using Radzen;

namespace BlazorDemos.Client.Components.Pages;

public partial class RadzenDemo
{
    [Inject]
    private DialogService DialogService { get; set; } = default!;

    private string Name { get; set; } = "";
    private string Message { get; set; } = "Enter a name and click the button.";

    private async Task ShowGreetingAsync()
    {
        if (string.IsNullOrWhiteSpace(Name))
        {
            Message = "Please enter a name.";
            return;
        }

        Message = $"Hello, {Name}!";

        // No JavaScript interop: the alert is a Radzen dialog rendered by Blazor.
        await DialogService.Alert($"Hello, {Name}!", "Greeting");
    }
}
