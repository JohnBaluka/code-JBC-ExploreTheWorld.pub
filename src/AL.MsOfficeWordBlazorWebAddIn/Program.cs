using ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client.Services;
using ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Components;

using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl;

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
builder.Services.AddScoped<WordDocumentInfo__Interop__Interface, WordDocumentInfo__Interop>();
builder.Services.AddScoped<WordEvents__Interop__Interface, WordEvents__Interop>();
builder.Services.AddScoped<WordSaveAsJson__Interop__Interface, WordSaveAsJson__Interop>();
builder.Services.AddScoped<WordCountriesExport__Interop__Interface, WordCountriesExport__Interop>();
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
    .AddAdditionalAssemblies(typeof(ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client._Imports).Assembly);

app.Run();
