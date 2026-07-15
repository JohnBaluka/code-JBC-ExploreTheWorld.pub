using BlazorDemos.Components;
using Radzen;

var builder = WebApplication.CreateBuilder(args);

var razorComponents = builder.Services.AddRazorComponents();
razorComponents.AddInteractiveServerComponents();
razorComponents.AddInteractiveWebAssemblyComponents();

// Registers Radzen services (DialogService, etc.) in the server/prerender scope.
builder.Services.AddRadzenComponents();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseAntiforgery();

app.MapStaticAssets();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddInteractiveWebAssemblyRenderMode()
    .AddAdditionalAssemblies(typeof(BlazorDemos.Client.Components.Pages.AllInOneDemo).Assembly);

app.Run();
