using ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Services;
using ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Components;

using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl;

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
builder.Services.AddScoped<ExcelWorkbookInfo__Interop__Interface, ExcelWorkbookInfo__Interop>();
builder.Services.AddScoped<ExcelEvents__Interop__Interface, ExcelEvents__Interop>();
builder.Services.AddScoped<ExcelSaveAsJson__Interop__Interface, ExcelSaveAsJson__Interop>();
builder.Services.AddScoped<ExcelCountriesExport__Interop__Interface, ExcelCountriesExport__Interop>();
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
    .AddAdditionalAssemblies(typeof(ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client._Imports).Assembly);

app.Run();
