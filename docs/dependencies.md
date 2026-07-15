# Dependency Tree

Project- and package-level dependencies for every project in the four `JBC.*.sln` solutions.
Generated from the `.csproj` files; keep it in sync when references change.

See also: [architecture.md](./architecture.md) (layer model + DB-manager pattern),
[project-templates.md](./project-templates.md) (`.csproj` shapes),
[shared-link-compilation.md](./shared-link-compilation.md) (how `._netF` twins link source).

## Overview

- **Layer rule:** `AL → BL → DL → CL` (never reversed). `CL` is the base; everything flows up.
- **`_netF` twins:** most core projects have a `net481` twin (`{Name}._netF`) that **shares the same source** via `<Compile Include>` links and mirrors the net10 project's references with down-level packages (EF Core 3.1.32, FluentAssertions 6.x, NetOfficeFw NuGet instead of the NetOffice project refs). Twins are listed compactly in [The `_netF` twins](#the-_netf-twins).
- **External references** (not in these solutions' `src/` tree):
  - **NetOffice** — the sibling repo `../code-zgh-NetOfficeFw__NetOffice__10/`: `NetOffice`, `NetOffice.Analyzers`, `OfficeApi`, `WordApi`, `ExcelApi`, `PowerPointApi`, `VBIDEApi` (net10 COM automation).
  - **Oqtane** — the `oqtane.framework/` submodule: `Oqtane.Client`, `Oqtane.Server`, `Oqtane.Shared`.
- **Decoupling invariants** (enforced after the AL/`_Impl` cleanup) are summarized in [Architectural invariants](#architectural-invariants).

## Legend

```
A → B      A has a ProjectReference to B
(pkg)      NuGet PackageReference
CL         core project        CL._netF   net481 twin
_Impl      interchangeable DL implementation (one per interface)
```

## The core spine

```
CL                     (no project references — pure primitives/helpers)
└─ DL                  → CL
   ├─ BL               → CL, DL, DL.CountriesNowSpaceApi
   │  └─ AL            → BL, CL, DL
   └─ (DL specialty + _Impl projects, below)
```

| Project | References | Notable packages |
|---------|-----------|------------------|
| `CL` | — | — |
| `DL` | CL | Microsoft.Extensions.Logging.Abstractions |
| `BL` | CL, DL, DL.CountriesNowSpaceApi | Microsoft.Extensions.DependencyInjection.Abstractions, …Logging.Abstractions |
| `AL` | BL, CL, DL | Microsoft.Extensions.Logging.Abstractions |

`DL` holds **all interfaces and shared data objects** (incl. the `MsOfficeApi` contracts:
`MsOfficeExportRepoFactory__Interface`, `MsOfficeDocument_Memory__Repo__Interface`,
`MsOfficeSaveAsJsonWriter__Interface`, the `Ms{Host}__Repo__Interface` set, and the
`CountriesNowSpaceApi_Interface` + row types). `BL` holds the concrete `{Name}__Service`
orchestrators. `AL` holds AL interfaces + framework-neutral app services.

## DL — specialty and implementation projects

Each interface in core `DL` is implemented by one or more `DL.*` projects. Interchangeable
implementations carry an `_Impl` suffix.

| Project | References | Notable packages |
|---------|-----------|------------------|
| `DL.CountriesNowSpaceApi` | CL, DL | — (HTTP `CountriesNowSpaceApi__Repo` + `FlagImageDownload__Repo`) |
| `DL.CountriesNowSpaceData` | CL, DL | Microsoft.EntityFrameworkCore(.Relational) 10.0.9 |
| `DL.CountriesNowSpaceData.SqliteDb_Impl` | DL.CountriesNowSpaceData | Microsoft.EntityFrameworkCore.Sqlite 10.0.9 |
| `DL.CountriesNowSpaceData.SqlServerDb_Impl` | DL.CountriesNowSpaceData | Microsoft.EntityFrameworkCore.SqlServer 10.0.9 |
| `DL.CountriesNowSpaceData.AccessDb_Impl` | DL.CountriesNowSpaceData | EntityFrameworkCore.Jet 10.0.1, System.Data.OleDb |
| `DL.CountriesNowSpaceData.InMemoryDb_Impl` | DL.CountriesNowSpaceData | Microsoft.EntityFrameworkCore.InMemory 10.0.9 |
| `DL.CountriesNowSpaceData.LocalStorageDb_Impl` | DL | Blazored.LocalStorage 4.0.1, …Components.Web |
| `DL.CountriesNowSpaceData.SessionStorageDb_Impl` | DL | Blazored.SessionStorage 2.4.0, …Components.Web |
| `DL.CountriesNowSpaceData.IndexedDb_Impl` | DL | Microsoft.JSInterop, …Components.Web |
| `DL.MsJSInterop` | CL, DL | Microsoft.JSInterop |
| `DL.MsJSInterop.RevealJs` | CL, DL, DL.MsJSInterop | Microsoft.JSInterop |
| `DL.MsOfficeApi.OpenXml_Impl` | CL, DL | DocumentFormat.OpenXml 3.5.1, System.Drawing.Common |
| `DL.MsOfficeApi.NetOffice_Impl` | CL, DL, **NetOffice, OfficeApi, WordApi, ExcelApi, PowerPointApi, NetOffice.Analyzers** | — |
| `DL.MsOfficeApi.Interop_Impl` | CL, DL | Microsoft.Office.Interop.{Word 15.0.4797, Excel 15.0.4795, PowerPoint 15.0.4420}, stdole |
| `DL.MsOfficeApi.Dynamic_Impl` | CL, DL | — (late-bound `dynamic` COM) |
| `DL.MsOfficeApi.Direct_Impl` | CL, DL | — (`MsOfficeDirectVbaRunner`, VBA `Application.Run`) |
| **`DL.MsOfficeApi_Impl`** (composition) | DL, **DL.MsOfficeApi.{OpenXml, NetOffice, Interop, Dynamic, Direct}_Impl**, NetOffice, OfficeApi, WordApi, ExcelApi, PowerPointApi | — (export factory + Save-As-JSON writers/dispatcher; the single project permitted to aggregate the five Office `_Impl`s) |
| `DL.MsOfficeApi.MsOfficeJs.Word_Impl` | DL, DL.MsJSInterop | Microsoft.JSInterop |
| `DL.MsOfficeApi.MsOfficeJs.Excel_Impl` | DL, DL.MsJSInterop | Microsoft.JSInterop |
| `DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl` | DL, DL.MsJSInterop | Microsoft.JSInterop |
| `DL.MsSystem` | CL, DL | Microsoft.Extensions.Logging.Abstractions |
| `DL.MsSystemNet` | CL, DL | — |

> **`DL.MsSystem` / `DL.MsSystemNet` have no source and no consumers.** They remain in the
> solutions but **no project references them** (see [Architectural invariants](#architectural-invariants)).

## AL — libraries

| Project | References | Notable packages |
|---------|-----------|------------------|
| `AL.BlazorLib` | AL, BL, CL, DL, DL.MsJSInterop, DL.MsJSInterop.RevealJs | Radzen.Blazor 11.1.2, …Components(.Web), …Localization |
| `AL.MauiLib` | AL.BlazorLib, AL, BL, CL, DL, DL.MsJSInterop, DL.MsJSInterop.RevealJs | Microsoft.Maui.Controls 10.0.80, …WebView.Maui, …Logging.Debug |
| `AL.WinFormsLib` | AL, AL.BlazorLib, BL, CL, DL, DL.CountriesNowSpaceApi, **NetOffice, OfficeApi, WordApi, ExcelApi, PowerPointApi, NetOffice.Analyzers** | …Components.WebView.WindowsForms |

> **The AL UI libraries reference no DL repo `_Impl` project.** `AL.WinFormsLib` keeps the
> **NetOffice** COM refs (its forms hold live NetOffice objects) but the five Office `_Impl`
> projects — plus the export factory, Save-As-JSON writers, and `MsOfficeSaveAsJsonWriter`
> dispatcher — live only in the **`DL.MsOfficeApi_Impl`** DL composition project (namespace
> `JBC.ExploreTheWorld.DL.MsOfficeApi_Impl`; listed in [DL — specialty and implementation
> projects](#dl--specialty-and-implementation-projects)), which the hosts reference. Forms reach it
> via the host-set `MsOfficeSaveAsJsonWriterProvider` / `MsOfficeExportRepoFactoryProvider` static seams.

## AL — hosts

**Blazor** (each browser host adds `OpenXml_Impl` for the in-memory `MsOfficeDocument_Memory__Repo`,
plus its DB providers):

| Host | References beyond `AL.BlazorLib, AL, BL, CL, DL` |
|------|--------------------------------------------------|
| `AL.BlazorWebApp` (server) | AL.BlazorWebApp.Client, DL.CountriesNowSpaceApi, **DL.MsOfficeApi.OpenXml_Impl**, DL.MsJSInterop, DL.CountriesNowSpaceData.{Sqlite,SqlServer,InMemory,Access}Db_Impl |
| `AL.BlazorWebApp.Client` (WASM) | DL.MsJSInterop, **DL.MsOfficeApi.OpenXml_Impl**, DL.CountriesNowSpaceData.{InMemory,LocalStorage,Session,Indexed}Db_Impl |
| `AL.BlazorWebApp.ClientOnly` (WASM PWA) | DL.CountriesNowSpaceApi, DL.MsJSInterop, **DL.MsOfficeApi.OpenXml_Impl**, DL.CountriesNowSpaceData.{InMemory,LocalStorage,Session,Indexed}Db_Impl |
| `AL.BlazorLib._radzen` (synced WASM app) | DL.CountriesNowSpaceApi, DL.MsJSInterop(.RevealJs), **DL.MsOfficeApi.OpenXml_Impl**, DL.CountriesNowSpaceData.{InMemory,LocalStorage,Session,Indexed}Db_Impl |
| `AL.BlazorLib.Server._radzen` | AL.BlazorLib._radzen, DL.CountriesNowSpaceApi, **DL.MsOfficeApi.OpenXml_Impl**, DL.MsJSInterop, DL.CountriesNowSpaceData.{Sqlite,SqlServer,InMemory,Access}Db_Impl |

**WinForms / MAUI / VSTO** (each references `DL.MsOfficeApi_Impl` — or its `._netF` twin —
for the export factory + writers, and sets the two provider seams at startup):

| Host | References |
|------|-----------|
| `AL.WinFormApp` | AL.WinFormsLib, **DL.MsOfficeApi_Impl**, DL.CountriesNowSpaceData.{Sqlite,SqlServer,InMemory,Access}Db_Impl, DL.MsJSInterop(.RevealJs) |
| `AL.MsOffice{Word,Excel,PowerPoint}VstoAddIn` | AL.WinFormsLib, **DL.MsOfficeApi_Impl**, DL.CountriesNowSpaceData(.Sqlite/Access)Db_Impl |
| `AL.MauiApp.WinUI` | AL.MauiLib, DL.CountriesNowSpaceApi, DL.CountriesNowSpaceData.{Sqlite,SqlServer,InMemory,Access}Db_Impl, **DL.MsOfficeApi_Impl** |
| `AL.MauiApp.{Droid,iOS,Mac}` | AL.MauiLib, DL.CountriesNowSpaceApi, DL.CountriesNowSpaceData.InMemoryDb_Impl |

> Each **MAUI head** supplies its own concrete DL repos (`CountriesNowSpaceApi__Repo`,
> `FlagImageStore_FileSystem__Repo`, `FlagImageDownload__Repo`) and references
> `DL.CountriesNowSpaceApi` itself — `AL.MauiLib` references no DL repo project.

**Console apps** (now share the composition factory instead of a hand-copied one):

| Host | References |
|------|-----------|
| `AL.ExportData.ConsoleApp` | BL, CL, DL, DL.CountriesNowSpaceApi, DL.CountriesNowSpaceData.{Sqlite,SqlServer,InMemory,Access}Db_Impl, **DL.MsOfficeApi_Impl** |
| `AL.SaveAsJson.ConsoleApp` | BL, CL, DL, **DL.MsOfficeApi_Impl** |

**Office web add-ins** (Blazor WASM; server host + `.Client` pair; the client references the
matching `DL.MsOfficeApi.MsOfficeJs.{Host}_Impl`):

| Host | References |
|------|-----------|
| `AL.MsOffice{Word,Excel,PowerPoint}BlazorWebAddIn` | its `.Client` (Radzen, …WebAssembly.Server) |
| `…BlazorWebAddIn.Client` | CL, DL.MsJSInterop, DL.MsOfficeApi.MsOfficeJs.{Host}_Impl (…WebAssembly, Radzen, Microsoft.JSInterop.WebAssembly, TypeScript.MSBuild) |

**Oqtane modules/theme** (reference the `oqtane.framework/` submodule projects):

| Project | References |
|---------|-----------|
| `AL.Oqtane.CountriesNow__Module.Server` | Oqtane.Shared, …Module.Client, DL.CountriesNowSpaceData.SqlServerDb_Impl, DL.CountriesNowSpaceApi, BL, AL.BlazorLib |
| `AL.Oqtane.CountrySlides__Module.Server` | Oqtane.Shared, …Module.Client, DL.CountriesNowSpaceData.SqlServerDb_Impl, DL.CountriesNowSpaceApi, DL.MsJSInterop.RevealJs, BL |
| `AL.Oqtane.{CountriesNow,CountrySlides}__Module.Client` | Oqtane.Shared, Oqtane.Client, AL.BlazorLib |
| `AL.Oqtane.Radzen__Module.{Client,Server}`, `AL.Oqtane.Theme` | Oqtane.Shared(/Client) (+ Radzen.Blazor) |

## Tests

| Test project | References (subjects) | Strategy / packages |
|--------------|----------------------|---------------------|
| `UnitTests` | CL, DL, DL.CountriesNowSpaceApi, BL, AL | Moq + FluentAssertions 8.x (mocks only) |
| `IntegrationTests` | CL, DL, DL.CountriesNowSpaceApi, DL.CountriesNowSpaceData, BL, AL | EF Core InMemory 10.0.9, Mvc.Testing |
| `OpenXmlLibTests` | CL, DL, DL.MsOfficeApi.OpenXml_Impl | real temp `.docx/.xlsx/.pptx` |
| `RazorTests` | AL.BlazorLib, AL, BL, CL, DL, DL.CountriesNowSpaceApi, DL.MsJSInterop | bUnit 2.7.2 + Moq |
| `WinFormAppTests` | DL, DL.MsOfficeApi.{Dynamic,NetOffice,OpenXml}_Impl | FlaUI (launches `AL.WinFormApp`) |
| `MauiAppTests`, `AccessDbTests` | — (launch exe/MSACCESS) | FlaUI |
| `OfficeAddinTests` | NetOffice, OfficeApi, WordApi, ExcelApi, PowerPointApi | FlaUI + NetOffice |
| `OfficeWebAddinTests`, `OqtaneTests`, `WebAppTests` | — (drive running host) | Playwright 1.61.0 |

## The `_netF` twins

Each net481 twin shares source with its net10 parent and mirrors its references with down-level
packages. Structural notes:

- Spine: `CL._netF ← DL._netF ← BL._netF ← AL._netF`; `DL.CountriesNowSpaceApi._netF`, `DL.CountriesNowSpaceData._netF` (+ `Sqlite/SqlServer/Access` twins), the five `DL.MsOfficeApi.*_Impl._netF`, `DL.MsSystemNet._netF`.
- `AL.WinFormsLib._netF` → CL._netF, DL._netF, BL._netF, DL.CountriesNowSpaceApi._netF, DL.CountriesNowSpaceData._netF (**no `_Impl`**); `DL.MsOfficeApi_Impl._netF` → DL._netF + the five `_Impl._netF`.
- Hosts: `AL.WinFormApp._netF` and the three `AL.MsOffice{…}VstoAddIn._netF` → `AL.WinFormsLib._netF` + `DL.MsOfficeApi_Impl._netF` + their DB `_Impl._netF`.
- Package down-levels: **EF Core 3.1.32** (Jet 3.1.1), **NetOfficeFw.* 1.9.10** NuGet (in place of the NetOffice project refs), **FluentAssertions 6.x** (tests), plus `DocumentFormat.OpenXml`, `Microsoft.CSharp`, and the Interop PIAs where used.

## External dependencies

| Group | Projects | Consumed by |
|-------|----------|-------------|
| **NetOffice** (`../code-zgh-NetOfficeFw__NetOffice__10/`) | NetOffice, NetOffice.Analyzers, OfficeApi, WordApi, ExcelApi, PowerPointApi, VBIDEApi | DL.MsOfficeApi.NetOffice_Impl, AL.WinFormsLib, DL.MsOfficeApi_Impl, OfficeAddinTests |
| **Oqtane** (`oqtane.framework/` submodule) | Oqtane.Client, Oqtane.Server, Oqtane.Shared | the `AL.Oqtane.*` modules/theme |

net481 hosts use the **NetOfficeFw.\*** NuGet packages (1.9.10) instead of these project refs.

## Solution membership

| Project group | `JBC.ExploreTheWorld` | `._netF` | `AL.BlazorLib._radzen` | `AL.BlazorWebApp` |
|---------------|:--:|:--:|:--:|:--:|
| Core (CL/DL/BL/AL) + `DL.*` | ✔ | ✔ (twins) | ✔ | ✔ |
| `AL.BlazorLib` + Blazor web hosts | ✔ | — | ✔ | ✔ |
| `_radzen` hosts | — | — | ✔ | — |
| WinForms / MAUI / VSTO (net10) + `MsOffice_Impl` | ✔ | — | — | — |
| WinForms / VSTO (net481) + `MsOffice_Impl._netF` | — | ✔ | — | — |
| Console apps, Office web add-ins | ✔ | — | — | — |
| Oqtane modules/theme + `Oqtane.*` | ✔ | — | — | ✔ |

## Architectural invariants

Enforced across the graph (verified by clean compiles of all four solutions):

1. **Layer direction** — `AL → BL → DL → CL`, never reversed. Tests reference their subjects only.
2. **AL UI libraries reference zero DL repo `_Impl` project.** `AL.BlazorLib`, `AL.MauiLib`,
   `AL.WinFormsLib`(/`._netF`) consume DL work through interfaces the host registers:
   - Browser export → `MsOfficeDocument_Memory__Repo__Interface` (host registers `OpenXml_Impl`).
   - WinForms export/watcher → the `DL.MsOfficeApi_Impl` composition project (the **single**
     lib holding the five Office `_Impl`s) reached via the provider seams.
3. **Composition/host owns the `_Impl` choice.** Hosts reference the concrete repo/factory projects
   and register them; the console apps and MAUI-WinUI head reuse the shared composition factory.
4. **`DL.MsSystem` / `DL.MsSystemNet` are unreferenced** (empty placeholders; kept in the solutions).
5. **`_netF` twins** never diverge structurally from their net10 parents — same source, down-level packages.
