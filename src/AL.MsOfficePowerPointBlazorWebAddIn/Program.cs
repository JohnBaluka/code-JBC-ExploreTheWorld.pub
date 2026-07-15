using ExploreTheWorld.AL.MsOfficePowerPointBlazorWebAddIn.Client.Services;
using ExploreTheWorld.AL.MsOfficePowerPointBlazorWebAddIn.Components;

using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl;

using Radzen;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddInteractiveWebAssemblyComponents();

builder.Services.AddRadzenComponents();
builder.Services.AddScoped<ServerCommandHandler>();

// JS interop — registered server-side too so server prerender of RenderModeInfo
// and the InteractiveWebAssembly pages can resolve the injected interfaces.
builder.Services.AddScoped<Layout__Interop__Interface, Layout__Interop>();
builder.Services.AddScoped<PowerPointPresentationInfo__Interop__Interface, PowerPointPresentationInfo__Interop>();
builder.Services.AddScoped<PowerPointEvents__Interop__Interface, PowerPointEvents__Interop>();
builder.Services.AddScoped<PowerPointSaveAsJson__Interop__Interface, PowerPointSaveAsJson__Interop>();
builder.Services.AddScoped<PowerPointCountriesExport__Interop__Interface, PowerPointCountriesExport__Interop>();
builder.Services.AddHttpClient();
builder.Services.AddScoped(sp => sp.GetRequiredService<IHttpClientFactory>().CreateClient());

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
    app.UseWebAssemblyDebugging();

    app.Lifetime.ApplicationStarted.Register(() =>
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = "/c npm run start-local",
            WorkingDirectory = app.Environment.ContentRootPath,
            UseShellExecute = false,
            CreateNoWindow = true
        });
    });
}
else
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseAntiforgery();
app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddInteractiveWebAssemblyRenderMode()
    .AddAdditionalAssemblies(typeof(ExploreTheWorld.AL.MsOfficePowerPointBlazorWebAddIn.Client._Imports).Assembly);

app.Run();
