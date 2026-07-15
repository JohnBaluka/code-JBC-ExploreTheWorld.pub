using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.AL;
using JBC.ExploreTheWorld.AL.BlazorWebApp;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;
using JBC.ExploreTheWorld.DL.MsJSInterop.FileDownload;
using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddInteractiveWebAssemblyComponents();

builder.Services.AddRadzenComponents();

builder.Services.AddScoped<RenderMode_AppService>();
builder.Services.AddSingleton(new Layout_AppService { AppTitle = "ETW BlazorWebApp" });

// No-op: New Window button is not meaningful during server-side pre-render.
// The WASM client (AL.BlazorWebApp.Client) registers BrowserNewWindow_AppService for interactive use.
builder.Services.AddScoped<NewWindow_AppService__Interface, NullNewWindow_AppService>();

// Browser OpenXML export — registered on server so static pre-render does not throw DI errors.
// The export button requires user interaction (interactivity); the pre-render pass never calls ExportAsync.
// The host supplies the OpenXML in-memory document builder (AL.BlazorLib references only the interface).
builder.Services.AddScoped<MsOfficeDocument_Memory__Repo__Interface, MsOfficeDocument_Memory__Repo>();
builder.Services.AddScoped<OfficeExport_AppService__Interface, BrowserExport_AppService>();

// WatcherEvent_AppService — registered as singleton. In the web app it is idle
// (no WinForms host pushing events), but Watcher pages require it via [Inject].
builder.Services.AddSingleton<WatcherEvent_AppService>();

// CountriesNow API client (HTTP)
builder.Services.AddHttpClient<CountriesNowSpaceApi_Interface, CountriesNowSpaceApi__Repo>(client =>
    client.BaseAddress = new Uri("https://countriesnow.space/api/v0.1/"));

// DbProvider_AppService — tracks the active provider; AvailableProviders drives the header dropdown.
var dbProviderName = builder.Configuration["DbProvider"] ?? DbProviderNames.InMemoryDb;
var dbProvider_AppService = new DbProvider_AppService
{
    ProviderName       = dbProviderName,
    AvailableProviders = [DbProviderNames.InMemoryDb, DbProviderNames.AccessDb, DbProviderNames.SqliteDb, DbProviderNames.SqlServerDb],
};
builder.Services.AddSingleton(dbProvider_AppService);

// Server DB providers — each registered as a keyed implementation; the switcher selects at call time.
builder.Services.AddExploreTheWorldInMemoryDbProvider();
builder.Services.AddExploreTheWorldAccessDbProvider(builder.Configuration.GetConnectionString("AccessDb"));
builder.Services.AddExploreTheWorldSqliteDbProvider(builder.Configuration.GetConnectionString("SqliteDb"));
builder.Services.AddExploreTheWorldSqlServerDbProvider(builder.Configuration.GetConnectionString("SqlServerDb"));

// Switcher — reads DbProvider_AppService.ProviderName at call time to select the active implementation.
builder.Services.AddCountriesNowSpaceDbSwitcher(() => dbProvider_AppService.ProviderName);

// BL manager
builder.Services.AddScoped<CountriesNowSpaceManager__Service>();

// Flag images — %LocalAppData% file cache with Wikimedia PNG download fallback.
builder.Services.AddSingleton<FlagImageStore__Repo__Interface, FlagImageStore_FileSystem__Repo>();
builder.Services.AddHttpClient<FlagImageDownload__Repo__Interface, FlagImageDownload__Repo>();
builder.Services.AddScoped<FlagImageManager__Service>();

// Generic JS interop (download-file.js, layout.js)
builder.Services.AddScoped<FileDownload__Interop__Interface, FileDownload__Interop>();
builder.Services.AddScoped<Layout__Interop__Interface, Layout__Interop>();

// Reveal.js JS interop for Country Slides page
builder.Services.AddScoped<RevealJs__Interop__Interface, RevealJs__Interop>();

var app = builder.Build();

// Ensure DB schema exists for every registered provider at startup.
// Unavailable providers (SQL Server not installed, Access engine missing) are silently skipped.
app.Services.TryEnsureExploreTheWorldInMemoryDbCreated();
app.Services.TryEnsureExploreTheWorldAccessDbCreated();
app.Services.TryEnsureExploreTheWorldSqliteDbCreated();
app.Services.TryEnsureExploreTheWorldSqlServerDbCreated();

if (app.Environment.IsDevelopment())
{
    app.UseWebAssemblyDebugging();
    app.UseDeveloperExceptionPage();
}
else
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.MapStaticAssets();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddInteractiveWebAssemblyRenderMode()
    .AddAdditionalAssemblies(
        typeof(JBC.ExploreTheWorld.AL.BlazorWebApp.Client.App).Assembly,
        typeof(Routes).Assembly);

app.Run();
