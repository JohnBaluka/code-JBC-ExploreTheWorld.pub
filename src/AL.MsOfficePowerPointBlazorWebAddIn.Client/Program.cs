using ExploreTheWorld.AL.MsOfficePowerPointBlazorWebAddIn.Client.Services;

using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl;

using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Radzen;

var builder = WebAssemblyHostBuilder.CreateDefault(args);

builder.Services.AddRadzenComponents();
builder.Services.AddScoped<ClientCommandHandler>();

// Generic layout interop (window breakpoint badge)
builder.Services.AddScoped<Layout__Interop__Interface, Layout__Interop>();

// PowerPoint Office.js page interops
builder.Services.AddScoped<PowerPointPresentationInfo__Interop__Interface, PowerPointPresentationInfo__Interop>();
builder.Services.AddScoped<PowerPointEvents__Interop__Interface, PowerPointEvents__Interop>();
builder.Services.AddScoped<PowerPointSaveAsJson__Interop__Interface, PowerPointSaveAsJson__Interop>();
builder.Services.AddScoped<PowerPointCountriesExport__Interop__Interface, PowerPointCountriesExport__Interop>();

// HttpClient for the CountriesNow page (direct calls to the countriesnow.space API from WASM)
builder.Services.AddScoped(sp => new HttpClient());

await builder.Build().RunAsync();
