using ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Services;

using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl;

using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Radzen;

var builder = WebAssemblyHostBuilder.CreateDefault(args);

builder.Services.AddRadzenComponents();
builder.Services.AddScoped<ClientCommandHandler>();

// Generic layout interop (window breakpoint badge)
builder.Services.AddScoped<Layout__Interop__Interface, Layout__Interop>();

// Excel Office.js page interops
builder.Services.AddScoped<ExcelWorkbookInfo__Interop__Interface, ExcelWorkbookInfo__Interop>();
builder.Services.AddScoped<ExcelEvents__Interop__Interface, ExcelEvents__Interop>();
builder.Services.AddScoped<ExcelSaveAsJson__Interop__Interface, ExcelSaveAsJson__Interop>();
builder.Services.AddScoped<ExcelCountriesExport__Interop__Interface, ExcelCountriesExport__Interop>();

// HttpClient for the CountriesNow page (direct calls to the countriesnow.space API from WASM)
builder.Services.AddScoped(sp => new HttpClient());

await builder.Build().RunAsync();
