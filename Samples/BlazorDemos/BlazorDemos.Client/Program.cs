using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Radzen;

var builder = WebAssemblyHostBuilder.CreateDefault(args);

// Registers Radzen's DialogService, NotificationService, TooltipService, etc.
builder.Services.AddRadzenComponents();

await builder.Build().RunAsync();
