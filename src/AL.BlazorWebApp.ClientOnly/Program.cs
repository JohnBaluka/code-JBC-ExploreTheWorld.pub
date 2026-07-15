using JBC.ExploreTheWorld.AL;
using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;
using JBC.ExploreTheWorld.DL.MsJSInterop.FileDownload;
using JBC.ExploreTheWorld.DL.MsJSInterop.FlagImageCache;
using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.InMemoryDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.LocalStorageDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SessionStorageDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.IndexedDb_Impl;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Radzen;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<Routes>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddRadzenComponents();

// CountriesNow external API (calls countriesnow.space directly from browser)
builder.Services.AddScoped<CountriesNowSpaceApi_Interface>(
    _ => new CountriesNowSpaceApi__Repo(new HttpClient
        { BaseAddress = new Uri("https://countriesnow.space/api/v0.1/") }));

// DB providers — each registered as a keyed implementation; the switcher selects at call time.
// InMemoryDb (EF Core in-memory) runs in WebAssembly alongside the browser-storage providers.
builder.Services.AddExploreTheWorldInMemoryDbProvider();
builder.Services.AddExploreTheWorldLocalStorageDbProvider();
builder.Services.AddExploreTheWorldIndexedDbProvider();
builder.Services.AddExploreTheWorldSessionStorageDbProvider();

// DbProvider_AppService — tracks the active provider; AvailableProviders drives the header dropdown.
var dbProvider_AppService = new DbProvider_AppService
{
    ProviderName       = DbProviderNames.InMemoryDb,
    AvailableProviders = [DbProviderNames.InMemoryDb, DbProviderNames.LocalStorageDb, DbProviderNames.IndexedDb, DbProviderNames.SessionStorageDb],
};
builder.Services.AddSingleton(dbProvider_AppService);

// Switcher — reads DbProvider_AppService.ProviderName at call time to select the active implementation.
builder.Services.AddCountriesNowSpaceDbSwitcher(() => dbProvider_AppService.ProviderName);

// BL manager
builder.Services.AddScoped<CountriesNowSpaceManager__Service>();

// Flag images — browser IndexedDB cache with Wikimedia PNG download fallback (CORS-enabled).
builder.Services.AddScoped<FlagImageCache__Interop__Interface, FlagImageCache__Interop>();
builder.Services.AddScoped<FlagImageStore__Repo__Interface, FlagImageStore_Browser__Repo>();
builder.Services.AddScoped<FlagImageDownload__Repo__Interface>(_ => new FlagImageDownload__Repo(new HttpClient()));
builder.Services.AddScoped<FlagImageManager__Service>();

// Generic JS interop (download-file.js, layout.js)
builder.Services.AddScoped<FileDownload__Interop__Interface, FileDownload__Interop>();
builder.Services.AddScoped<Layout__Interop__Interface, Layout__Interop>();

// Reveal.js JS interop for Country Slides page
builder.Services.AddScoped<RevealJs__Interop__Interface, RevealJs__Interop>();

// Browser OpenXML export (generates in-memory, downloads via JS)
// The host supplies the OpenXML in-memory document builder (AL.BlazorLib references only the interface).
builder.Services.AddScoped<MsOfficeDocument_Memory__Repo__Interface, MsOfficeDocument_Memory__Repo>();
builder.Services.AddScoped<OfficeExport_AppService__Interface, BrowserExport_AppService>();

builder.Services.AddScoped<RenderMode_AppService>();
builder.Services.AddSingleton(new Layout_AppService { AppTitle = "ETW BlazorWebApp.ClientOnly" });

// WatcherEvent_AppService is idle in WASM but required by Watcher pages via [Inject]
builder.Services.AddSingleton<WatcherEvent_AppService>();

// NewWindow_AppService__Interface — opens the app in a new browser tab.
builder.Services.AddScoped<NewWindow_AppService__Interface, BrowserNewWindow_AppService>();

await builder.Build().RunAsync();
