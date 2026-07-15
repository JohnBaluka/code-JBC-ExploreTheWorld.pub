using ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client.Services;

using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl;

using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Radzen;

var builder = WebAssemblyHostBuilder.CreateDefault(args);

builder.Services.AddRadzenComponents();
builder.Services.AddScoped<ClientCommandHandler>();

// Generic layout interop (window breakpoint badge)
builder.Services.AddScoped<Layout__Interop__Interface, Layout__Interop>();

// Word Office.js page interops
builder.Services.AddScoped<WordDocumentInfo__Interop__Interface, WordDocumentInfo__Interop>();
builder.Services.AddScoped<WordEvents__Interop__Interface, WordEvents__Interop>();
builder.Services.AddScoped<WordSaveAsJson__Interop__Interface, WordSaveAsJson__Interop>();
builder.Services.AddScoped<WordCountriesExport__Interop__Interface, WordCountriesExport__Interop>();

// HttpClient for the CountriesNow page (direct calls to the countriesnow.space API from WASM)
builder.Services.AddScoped(sp => new HttpClient());

await builder.Build().RunAsync();
