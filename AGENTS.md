# AGENTS.md — code-JBC-ExploreTheWorld

This file is the canonical AI coding guide for **ExploreTheWorld**.

For human-facing product context, see [README.md](./README.md).
Do not treat README.md as coding instructions unless this file explicitly says so.

## Project Identity

- Repo: `code-JBC-ExploreTheWorld`
- Owner namespace: `JBC`
- Purpose: REST API wrapper application demonstrating JBC layered architecture patterns. Supports both .NET 10.0 (primary) and .NET Framework 4.8.1 via shared-link compilation. Includes Blazor WebAssembly (hybrid SSR+WASM and standalone ClientOnly), WinForms (BlazorWebView via `AL.WinFormsLib`), a .NET MAUI multi-project app (`AL.MauiLib` shared library + `AL.MauiApp.{WinUI|Droid|iOS|Mac}` heads), and Office Add-in support.

## Read First

1. [README.md](./README.md) — project overview
2. [docs/architecture.md](./docs/architecture.md) — layered architecture (CL/DL/BL/AL), dependency rules, DB-backed manager pattern
3. [docs/naming-conventions.md](./docs/naming-conventions.md) — file naming, namespace patterns, class/method naming rules
4. [docs/file-structure.md](./docs/file-structure.md) — folder organization by layer and specialty
5. [docs/project-templates.md](./docs/project-templates.md) — `.csproj` SDK types, `<PropertyGroup>` patterns, package references by layer
6. [docs/shared-link-compilation.md](./docs/shared-link-compilation.md) — how `.NET Framework` projects link source from `.NET 10` projects
7. [docs/dependencies.md](./docs/dependencies.md) — project- and package-level dependency tree for the four `JBC.*.sln` solutions

## Coding Rules

- Dependency rule: AL → BL → DL → CL. Never reverse.
- DL references must not be injected into or instantiated by AL **UI items** (pages, components, forms). UI items inject **BL** classes (or AL service interfaces); DL is injected into BL. Add new BL orchestrator classes where needed (e.g. `MsOfficeExportManager__Service`).
- **BL services** are the business-logic orchestrators (`{Name}__Service`) and live in `BL/_Services/` (namespace `JBC.ExploreTheWorld.BL` — a leading-underscore folder is dropped from the namespace; see [docs/naming-conventions.md](./docs/naming-conventions.md)). A BL service is **concrete — it has no interface of its own**: its dependencies are DL repo interfaces, so unit tests exercise the real service and mock those DL interfaces (a sibling BL service is likewise built from its own mocked DL deps). **BL defines no interfaces.** The seam that lets a **host/AL layer** supply the concrete platform-specific export repos is a **DL contract**, `DL.MsOfficeApi.MsOfficeExportRepoFactory__Interface` (implemented by `MsOfficeExportRepoFactory` in the `DL.MsOfficeApi_Impl` composition project, or by each console host's own copy), so BL consumes it without ever referencing the DL `_Impl` projects.
- All **DL** interfaces and shared data objects live in the core `ExploreTheWorld.DL` project; the `DL.*` projects only implement them. **Exception:** a BL composite/switcher service may implement a DL repo interface to orchestrate multiple DL implementations at runtime (e.g. `DbProviderSwitcher__Service`, which selects among the keyed `CountriesNowSpaceData.*Db_Impl` providers; cf. `MsOfficeExportManager__Service`). These composites live in **BL**, never in a `DL.*` project.
- Interchangeable DL implementations of one repo interface are grouped under a shared parent namespace with an `_Impl` suffix per implementation: the Office repos are `DL.MsOfficeApi.{Interop|NetOffice|OpenXml|Dynamic|Direct}_Impl` (contracts in `DL.MsOfficeApi`; `Direct_Impl` is JSON-write only — its `ExportAsync` throws `NotSupportedException`); the DB providers are `DL.CountriesNowSpaceData.{Sqlite|SqlServer|Access|InMemory|LocalStorage|SessionStorage|Indexed}Db_Impl`. Provider-key strings are centralized in `DbProviderNames` (core DL).
- `Interop_Impl` (net10) + `Interop_Impl._netF` (net481 shared-link twin) mirror `NetOffice_Impl` but use the `Microsoft.Office.Interop.{Word|Excel|PowerPoint}` PIAs (NuGet 15.x) instead of NetOffice. `Microsoft.Office.Core` (office.dll, v15.0.0.0) has no NuGet package, so the `.csproj` references the matching PIA already on any Office/VSTO dev machine (GAC first, VSTO Shared PIA folder fallback). Newer Office members absent from the Office15 PIA (`Shape.Decorative`/`IsNarration`, `Presentation.AutoSaveOn`/`ReadOnlyRecommended`) are read late-bound via `((dynamic)x).Member` so `ComTryGet` still swallows them to null. Register with `services.AddInteropLib()`.
- **Wherever `NetOffice` is a user-selectable Office-API option, `Interop` is too.** MsOfficeApi picklist order, where the members apply: **COM, Direct, Interop, Dynamic, NetOffice, OpenXML** (COM has no implementation yet; Direct is Save-As-JSON only, not an Export-Data method). The `JsonWriteMethod`/`ExportMethod` enums and every method dropdown/factory follow this order.
- DB-provider picklist order (`AvailableProviders`), where the members apply: **InMemoryDb, AccessDb, SqliteDb, SqlServerDb, LocalStorageDb, IndexedDb, SessionStorageDb**. `InMemoryDb` is offered on every host that has a provider mechanism (server, MAUI WinUI, and — via EF Core in-memory — the WASM browser hosts). net481-only hosts (`AL.WinFormApp._netF`, Oqtane) have no in-memory provider (`InMemoryDb_Impl` is net10-only) and no provider picklist.
- **`InMemoryDb` is the default-selected provider on every host where it is available** (the switchable net10 hosts — `AL.BlazorWebApp`, `AL.BlazorWebApp.Client`, `AL.BlazorWebApp.ClientOnly`, `AL.MauiApp.WinUI`, `AL.WinFormApp`, `AL.ExportData.ConsoleApp`, plus the InMemory-only MAUI mobile heads). It requires no external database, so the app runs out of the box; other providers stay user-selectable at runtime. Seed it via `config["DbProvider"] ?? DbProviderNames.InMemoryDb` (and `"DbProvider": "InMemoryDb"` in `appsettings.json`). Hosts where InMemory is not registered (VSTO add-ins → `SqliteDb`; Oqtane → `SqlServerDb`) keep their own default.
- **When opening/automating an Office application, set `Visible = true`** (`Application.Visible = MsoTriState.msoTrue` for PowerPoint; `-1` late-bound). This applies to both the Export-Data (`ExportAsync`) and the Save-As-JSON reader (`WriteDocumentJsonAsync`) paths across `NetOffice_Impl`/`Interop_Impl`/`Dynamic_Impl`/`Direct_Impl`, and to the Access/Office VBA writers. PowerPoint rejects `Visible = msoFalse`; open its presentation `WithWindow=msoTrue` so it surfaces. Attach-only paths (WebView watcher forms, `MauiWatcher_AppService`) that connect to an already-running instance leave `Visible` as-is.
- Default Office file names embed the API choice: Export-Data documents are `{Base}[-Access]-{Method}.{ext}` (`-Access` when the active DB provider is `AccessDb`) via `CL.MsOfficeExportName_Helper`; Save-As-JSON output inserts the method before the extension via `SaveAsJson_Helper.BuildDefaultPath(current, fallback, method)` (e.g. `Report-Direct.pptx.json`), recomputed whenever the active document **or the selected method** changes — the Blazor watcher components wire the method dropdown's `Change` to `OnJsonMethodChanged`; the `._netF` watcher forms wire `cbxJsonWriteMethod.SelectedIndexChanged` to `UpdateDefaultOutputPath()` (a user-customized path is left untouched).
- All **AL** interfaces and framework-neutral data objects live in the core `ExploreTheWorld.AL` project; the `AL.*` projects only implement them. (Blazor/WinForms-coupled service *implementations* stay in their `AL.*` lib.)
- **AL UI libraries reference no DL repo `_Impl` project.** They consume DL work through interfaces registered by the host, never by instantiating a `_Impl` repo. Two concrete seams enforce this: (1) the browser in-memory document builder is a DL contract `DL.MsOfficeApi.MsOfficeDocument_Memory__Repo__Interface` (implemented by `OpenXml_Impl`, injected into `AL.BlazorLib`'s `BrowserExport_AppService`, registered by each browser host); (2) the WinForms Office export factory + Save-As-JSON writers live in the **`DL.MsOfficeApi_Impl` (+`._netF`) composition project** (the single lib permitted to reference the five `DL.MsOfficeApi.*_Impl` projects), which the WinForms/VSTO/MAUI-WinUI hosts reference. Because the net481 watcher forms have no DI container, `AL.WinFormsLib` exposes two host-set static seams — `MsOfficeSaveAsJsonWriterProvider` and `MsOfficeExportRepoFactoryProvider` — that each host assigns once at startup to the composition project's concrete types.
- **AL application/platform services use the `{Name}_AppService` suffix** (single underscore + `App`), e.g. `Layout_AppService`, `BrowserExport_AppService`, `MauiWatcher_AppService`. The `App` prefix on `Service` deliberately distinguishes an **AL host-plumbing service** from a **BL `{Name}__Service`** orchestrator, so the two never blur together. **Every `Service`-named type in an `AL*` project uses `_AppService`** — never a bare `_Service`/`Service` suffix. When several hosts supply platform-specific implementations of one seam, the contract is `{Name}_AppService__Interface` in core `ExploreTheWorld.AL` (`NewWindow_AppService__Interface`, `OfficeExport_AppService__Interface`), implemented per host (`BrowserNewWindow_AppService`, `WinFormNewWindow_AppService`, `MauiNewWindow_AppService`, `NullNewWindow_AppService`). Third-party service **types** keep their own names (`Radzen.DialogService`, `IServiceCollection`, `ServiceCollectionExtensions`) — we do not rename them — **but the injected field/property/variable that holds a service uses the `_AppService` suffix on its own role-based base name**, ours and third-party alike (`Dialog_AppService`, `Notification_AppService`, `Export_AppService`, `Watcher_AppService`, local `dbProvider_AppService`); keep the member's base, not the full type name. A member may end up matching its type name (`Layout_AppService Layout_AppService`) — legal C#. See [docs/naming-conventions.md](./docs/naming-conventions.md) → Service Naming.
- DB-backed manager pattern: check-DB → API-fallback → persist → return `DataResult_Row<T>` with `DataSource_Enum`.
- DL repo pair naming: `{Name}Manager__Repo__Interface` (interface, in core DL) + `{Name}Manager__Repo` (implementation, in a `DL.*` project). BL orchestrators use `{Name}__Service` instead (concrete, no paired interface — see the BL-services rule above).
- Interface naming: `_Interface` suffix (not `I` prefix). DL repo interfaces: `__Repo__Interface` suffix.
- **Leading-underscore folders (`_Services/`, `_Export/`, `_Forms/`, `_Shared/`, …) do not contribute a namespace segment** — files in them take the namespace of the nearest non-underscore ancestor folder.
- **Group by type, not feature, and only when 2+.** Repos and their repo interfaces go in `_Repos/`, factories in `_Factories/` (class suffix `__Factory`), entities in `_Entities/`, rows in `_Rows/`, EF field constants in `_Fields/`. No per-feature folders (`Managers/`, `CountriesNowSpaceApiManager/`, `MsOfficeDocumentManager/`). A lone repo/factory sits at the project root (no `_Repos`/`_Factories` folder). So the DB-manager repo interface + `cns_*` entities live at the core `DL` root (`JBC.ExploreTheWorld.DL`); each `*Db_Impl` keeps its single repo + `ExploreTheWorldDbContext__{Provider}__Factory` at its root; MsOffice `*_Impl` repos sit in `_Repos/`.
- Namespace pattern: `JBC.ExploreTheWorld[.{Layer}][.{Specialty}][.{Domain}]`
- Method naming: `Get{Entity}`, `Create{Entity}`, `{Action}Async` for async, `Is{Condition}` / `Has{Feature}` for booleans.
- One type per file: in all `ExploreTheWorld.*` projects, every class, enum, interface, record, and struct (public or internal) lives in its own `.cs` file named after the type. Razor code-behind files (`Component.razor.cs`) contain only that component's partial class. Third-party content (`oqtane.framework/`, `node_modules/`) is exempt.
- Do not abbreviate: HTTP→Http, API→Api, JSON→Json, XML→Xml, HTML→Html, UI→UserInterface.
- Use abbreviations: GUID, ID, URL, JS, CSS.
- "DL" = **Dependency Layer**, not "Data Layer" (DB, HTTP, system, and JS-interop dependencies wrapped behind interfaces).
- JS interop belongs in a dedicated interop project, never inline in `.razor`: generic/reveal.js interop lives in `DL.MsJSInterop[.RevealJs]`; per-host Office.js page interop lives in `DL.MsOfficeApi.MsOfficeJs.{Word|Excel|PowerPoint}_Impl` (grouped under `MsOfficeApi` with the `_Impl` suffix, but still referencing `DL.MsJSInterop` for `JsModuleInterop__Base`). Add a `{Name}__Interop` (deriving `JsModuleInterop__Base`) + `{Name}__Interop__Interface`, register it in `Program.cs` (`AddScoped<…__Interface, …>()`), and inject the interface. Do not `IJSRuntime.InvokeAsync("import", …)` or define `window.*` globals in components/`App.razor`. (Office `[JSImport]`/`JSHost` `SharedUtils.js` and `commands.js`/`*.lib.module.js` add-in infrastructure are the documented exceptions.)

## Build and Test

All `.sln` files and .NET project folders live under `src/`; test projects live under `src/Tests/`.

**Node.js is not a build prerequisite.** The six Office web add-in projects run an `NpmRestore` target for dev-only tooling (add-in debugging/manifest scripts, Office.js editor typings); it probes for npm and skips with a warning when Node.js is absent, so `JBC.ExploreTheWorld.sln` builds without it. Never let that target hard-fail the build — see [docs/project-templates.md](./docs/project-templates.md) → npm Restore Target.

**Launching a web add-in** needs a sideloaded manifest, a running server, and the Office app started together — `Scripts\Start-WebAddin.ps1` does all three without Node.js (`npm run start-local` remains the Node equivalent). See [docs/project-templates.md](./docs/project-templates.md) → Launching a web add-in, including the known limitation that the `ETW (Web)` ribbon tab does not appear on Microsoft 365 build 16.0.20228.

```powershell
# All projects
dotnet build src\JBC.ExploreTheWorld.sln

# .NET Framework only (always run after shared code changes)
dotnet build src\JBC.ExploreTheWorld._netF.sln

# Blazor app
dotnet build src\JBC.ExploreTheWorld.AL.BlazorLib._radzen.sln

# Blazor ClientOnly standalone (WASM-only PWA)
dotnet build src\AL.BlazorWebApp.ClientOnly\ExploreTheWorld.AL.BlazorWebApp.ClientOnly.csproj

# Run all test projects in the prescribed order
.\Scripts\Run-AllTests.ps1

# Launch an Office web add-in: sideload manifest + start server + open the Office app
.\Scripts\Start-WebAddin.ps1 -OfficeApp Word          # or Excel / PowerPoint
.\Scripts\Start-WebAddin.ps1 -OfficeApp Word -Unregister

# Run .NET 10 tests
dotnet test src\Tests\UnitTests\ExploreTheWorld.UnitTests.csproj
dotnet test src\Tests\IntegrationTests\ExploreTheWorld.IntegrationTests.csproj
dotnet test src\Tests\OpenXmlLibTests\ExploreTheWorld.OpenXmlLibTests.csproj
dotnet test src\Tests\RazorTests\ExploreTheWorld.RazorTests.csproj
dotnet test src\Tests\WebAppTests\ExploreTheWorld.WebAppTests.csproj
dotnet test src\Tests\WinFormAppTests\ExploreTheWorld.WinFormAppTests.csproj
dotnet test src\Tests\MauiAppTests\ExploreTheWorld.MauiAppTests.csproj
dotnet test src\Tests\OfficeAddinTests\ExploreTheWorld.OfficeAddinTests.csproj
dotnet test src\Tests\OfficeWebAddinTests\ExploreTheWorld.OfficeWebAddinTests.csproj
dotnet test src\Tests\OqtaneTests\ExploreTheWorld.OqtaneTests.csproj
dotnet test src\Tests\AccessDbTests\ExploreTheWorld.AccessDbTests.csproj

# Run .NET Framework tests
dotnet test src\Tests\UnitTests._netF\ExploreTheWorld.UnitTests._netF.csproj
dotnet test src\Tests\IntegrationTests._netF\ExploreTheWorld.IntegrationTests._netF.csproj
dotnet test src\Tests\SqliteDbTests._netF\ExploreTheWorld.SqliteDbTests._netF.csproj
dotnet test src\Tests\OpenXmlLibTests._netF\ExploreTheWorld.OpenXmlLibTests._netF.csproj
dotnet test src\Tests\WinFormAppTests._netF\ExploreTheWorld.WinFormAppTests._netF.csproj
dotnet test src\Tests\OfficeAddinTests._netF\ExploreTheWorld.OfficeAddinTests._netF.csproj
```

## Test Project Conventions

| Project | Strategy | Notes |
|---------|----------|-------|
| `UnitTests` | net10.0 · mocks only | No data access; uses Moq + FluentAssertions 8.x |
| `UnitTests._netF` | net481 · shared-link from `UnitTests` | Links all source; FluentAssertions 6.x |
| `IntegrationTests` | net10.0 · EF Core InMemory | `IDbContextFactory<T>` fixture; FluentAssertions 8.x |
| `IntegrationTests._netF` | net481 · EF Core 3.x InMemory · standalone | `TestDbContextFactory` wraps `IDbContextFactory<T>` (EF Core 3.x — `CreateDbContext()` only); FluentAssertions 6.x |
| `OpenXmlLibTests` | net10.0 · real temp files | Exports `.docx`/`.xlsx`/`.pptx` via `DL.MsOfficeApi.OpenXml_Impl`, re-opens and verifies; FluentAssertions 8.x |
| `OpenXmlLibTests._netF` | net481 · shared-link from `OpenXmlLibTests` | Links all source; tests `DL.MsOfficeApi.OpenXml_Impl._netF`; FluentAssertions 6.x |
| `SqliteDbTests._netF` | net481 · EF Core 3.x Sqlite · standalone | Real SQLite temp file; no net10.0 counterpart (InMemory covered by `IntegrationTests`) |
| `WinFormAppTests` | net10.0-windows · FlaUI | Tests `AL.WinFormApp` (BlazorWebView exe); forms/services live in `AL.WinFormsLib`. `JsonWriters/` compares the NetOffice/Dynamic/OpenXml Save-As-JSON writers (returns early when PowerPoint is not installed) |
| `WinFormAppTests._netF` | net481 · FlaUI · standalone | Tests `AL.WinFormApp._netF`; forms/services live in `AL.WinFormsLib._netF` |
| `MauiAppTests` | net10.0-windows · FlaUI | Tests `AL.MauiApp.WinUI` (unpackaged MAUI Blazor exe); the Droid/iOS/Mac heads are InMemoryDb-only and not UI-tested |
| `OfficeAddinTests` | net10.0-windows · FlaUI + NetOffice project refs | Tests VSTO add-in ribbon via UI automation |
| `OfficeAddinTests._netF` | net481 · FlaUI + NetOfficeFw NuGet · shared-link from `OfficeAddinTests` | Links all source; uses `NetOfficeFw.*` packages |
| `OqtaneTests` | net10.0 · Playwright | Tests Oqtane modules + ETW theme against the local Oqtane host (LocalDB site); fixture starts `Oqtane.Server.exe` if not already listening |
| `AccessDbTests` | net10.0-windows · FlaUI | Launches MSACCESS.EXE with `VBA\Access\ExploreTheWorld.accdb` read-only (`/ro`); skips when Access is not installed |

UI test bases foreground the target window and wait `ETW_TEST_SETTLE_MS` (defaults 1500–4000 ms per project) before screenshots — WebView2/Blazor paints asynchronously, so capturing immediately produces blank images. In Playwright test projects, navigate with the base-class `NavigateAsync(url)` (never `Page.GotoAsync`) so the "before" screenshot is taken after the page has painted.

## WinFormsLib Projects

`AL.WinFormsLib` (net10.0-windows, Razor SDK) and `AL.WinFormsLib._netF` (net481) are Windows Forms Control Libraries that share all WinForms UI code (Forms, UserControls, export/watcher helpers) between `AL.WinFormApp`, `AL.WinFormApp._netF`, and the six VSTO add-in projects. Add-ins reference these libraries via `ProjectReference` instead of `<Compile Include>` file-links.

- `AL.WinFormsLib` — net10 forms: `ExploreTheWorld_Form`, `CountriesNowSpace_WebView_Form`, watcher `*_WebView_Form`s, export helpers (`_Export/`), watcher helpers (`_Watcher/`), and net10-only services (`_Services/`).
- `AL.WinFormsLib._netF` — net481 native forms: `Main_Form`, `CountriesNowSpace_Form`, watcher `*_Form`/`*_UserControl`s. Share-links the portable `_Export/` and `_Watcher/` helper files from `AL.WinFormsLib` (shared-link compilation pattern).
- Namespace for both: `JBC.ExploreTheWorld.AL.WinFormsLib`.
- **Neither UI library references a DL `_Impl` project.** The Office export factory (`MsOfficeExportRepoFactory`), the Save-As-JSON writer wrappers (`Ms{Word|Excel|PowerPoint}_JsonWriter`, `MsOfficeJsonWriter_Helper`), and the writer dispatcher (`MsOfficeSaveAsJsonWriter`) live in the **`DL.MsOfficeApi_Impl`** (net10-windows) / **`DL.MsOfficeApi_Impl._netF`** (net481, shared-linked) composition project — the single lib that references `DL.MsOfficeApi.{OpenXml|NetOffice|Interop|Dynamic|Direct}_Impl` (namespace `JBC.ExploreTheWorld.DL.MsOfficeApi_Impl`). The UI forms reach these through the host-set static providers `MsOfficeSaveAsJsonWriterProvider` / `MsOfficeExportRepoFactoryProvider` (in `AL.WinFormsLib`, consuming the core-DL contracts). **Every WinForms/VSTO host references the composition project and sets both providers at startup** (`WinFormApp`, `WinFormApp._netF`, the six VSTO add-ins). The MAUI WinUI head references the composition project too and sets the writer provider (its export uses the DI-registered factory).

## MauiApp Projects

The MAUI app uses the **multi-project** layout: `AL.MauiLib` (net10.0 shared library — `App`, `MainPage`, `MauiProgramExtensions.UseSharedMauiLib()` with the platform-neutral registrations) plus one head per platform: `AL.MauiApp.WinUI`, `AL.MauiApp.Droid`, `AL.MauiApp.iOS`, `AL.MauiApp.Mac`. The shared project is a `Lib` (not `App`) because it targets no single platform — mirroring the `AL.WinFormsLib` / `AL.WinFormApp` split.

- `UseSharedMauiLib()` registers only the platform-neutral **BL orchestrators** (`CountriesNowSpaceManager__Service`, `FlagImageManager__Service`) and JS interops; `AL.MauiLib` references **no** DL repo project. Each head supplies the concrete DL repos it depends on — `CountriesNowSpaceApi__Repo`, `FlagImageStore_FileSystem__Repo`, `FlagImageDownload__Repo` — and references `DL.CountriesNowSpaceApi` itself.

- Only the **WinUI** head supports Watchers, Export API Data, and the four switchable server EF providers (Sqlite/SqlServer/Access/InMemory via the BL `DbProviderSwitcher__Service`); the other heads register `AddExploreTheWorldInMemoryDb()` only and set `Layout_AppService { ShowWatcherNavItems = false, ShowExportOptions = false }`.
- Each head owns its `wwwroot/index.html`, platform bootstrap files, and `MauiProgram`; the shared project must stay platform-neutral (no server EF provider / NetOffice / Dynamic references, no `Platforms/` folder).
- The shared project references `Microsoft.AspNetCore.Components.WebView.Maui` with `ExcludeAssets="build;buildTransitive"`; heads reference it normally (both referencing build assets duplicates static web assets).

## Reusable Skills

Apply relevant guidance from:
- [../code-JBC/skills/repo-conventions.md](../code-JBC/skills/repo-conventions.md) — repo naming, path patterns, AssemblySuffix rules
- [../code-JBC/skills/dotnet-blazor.md](../code-JBC/skills/dotnet-blazor.md) — layered architecture, Blazor patterns, shared-link compilation
- [../code-JBC/skills/git-workflow.md](../code-JBC/skills/git-workflow.md) — cross-repo dependency model, build order

## Razor Attribute Formatting

For **all elements** in `.razor` files (HTML elements and Blazor components alike):

- **First attribute stays on the same line** as the opening tag; every additional attribute goes on its own line, aligned under the first attribute.
- **Attribute order** — always group 1 → 2 → 3:
  1. **Native HTML attributes** — `id`, `class`, `style` (lowercase) must come first; remaining native HTML attrs follow in their existing order.
  2. **Blazor binding attributes** (prefixed with `@`) — `@ref`, `@attributes` must come first; other `@` attrs follow in their existing order.
  3. **Component-specific parameters** (capitalized Blazor params) — `Style` (capital S) must come first; other component params follow in their existing order.
- **When attributes are split across lines**, text content and the closing tag each go on their own line. The content is indented to the continuation-attribute column (opening-tag base indent + `<TagName `.length). The closing tag returns to the opening tag's base indent.
- Do **not** modify files under `oqtane.framework/` (third-party submodule).

```razor
@* Single attribute — content stays inline *@
<RadzenText TextStyle="TextStyle.H5">Not Found</RadzenText>

@* Multiple attributes — content and closing tag each on their own line *@
<RadzenText Style="color:red;"
            TextStyle="TextStyle.H3">
            Error.
</RadzenText>

@* HTML element example — attributes *@
<div id="myId"
     class="my-class"
     style="color:red"
     @ref="myRef"
     @attributes="Attributes">

@* Blazor component example — attributes *@
<RadzenStack style="height:100%"
             @ref="myRef"
             Style="padding:1rem"
             Orientation="Orientation.Vertical"
             Gap="1rem"
             Class="p-4">
```

## UI Styling

- **Radzen only** — `Microsoft.FluentUI.AspNetCore.Components` is not referenced and must not be added.
- Radzen CSS file: `_content/Radzen.Blazor/css/default.css` — this is the only complete free theme (includes `@font-face` for icons + all base styles). `fluent.css` / `fluent-base.css` are overrides-only or require a paid Radzen Studio subscription and must NOT be used.
- CSS variables: use Radzen tokens (`--rz-base-900`, `--rz-primary`, `--rz-text-contrast-color`, `--rz-body-background-color`). Never use Fluent UI tokens (`--neutral-foreground-rest`, `--accent-fill-rest`, `--fill-color`, `--type-ramp-*`, etc.).
- Do not create scoped `*.razor.css` files that duplicate global layout rules already in `app.css` — Blazor's scoped attribute selector has higher specificity and silently overrides global styles.
- Office add-in task pane layout: use `RadzenLayout` with **no `RadzenSidebar`** — the pane is too narrow (~350 px). Place `<NavMenu />` at the top of `RadzenBody` instead.
- Task pane nav uses `RadzenMenu` (horizontal) with short labels; `RadzenPanelMenu` is for full-width sidebars only.
- Inline checkbox/label pairs: wrap each in `<RadzenStack Orientation="Orientation.Horizontal" AlignItems="AlignItems.Center">` — a vertical `RadzenStack` puts them on separate rows.

## Branding / Display Name

- **Title bars show a short `ETW <ProjectName>` name.** For every app **except the Access database and the VBA macro add-ins**, the window / browser-tab / task-pane title bar shows the project name with the `ExploreTheWorld.AL.` prefix shortened to `ETW ` — e.g. `ETW BlazorWebApp`, `ETW WinFormApp`, `ETW MsOfficeWordVstoAddIn`. `._netF` variants keep the suffix (`ETW WinFormApp._netF`). This is wired as:
  - **Blazor tab title** — each host sets `Layout_AppService.AppTitle` to its project name; `Main_Layout` renders it as the single `<PageTitle>` (shared pages no longer set their own). Browser hosts also set the static `<title>` in their host HTML (`App.razor` / `index.html`).
  - **WinForms** — `ExploreTheWorld_Form` / `Main_Form` `Text` (Designer).
  - **MAUI** — per-head `App.AppTitle` (set in each head's `MauiProgram`) drives `Window.Title` / `MainPage.Title` / `MauiNewWindow_AppService`.
  - **VSTO net10** — the floating WebView `form.Text` set in each `Addin.cs`.
  - **VSTO net481** — the two `[CustomPane(...)]` titles carry the project name plus a function suffix so the panes stay distinguishable (`… — Watcher`, `… — Countries API`).
  - **Web add-ins** — the manifest `<DisplayName>` (the Office pane caption; also the install-UI name).
- The **user-facing brand name "Explore the World"** (three words, lowercase `the`) stays as **in-body** text only: header/brand headings (Radzen header `H6`, Home `H4`, web-add-in `MainLayout` header), the PowerPoint export title slide, Office add-in manifest `Description`/GetStarted callouts, the VSTO registry `FriendlyName`/`Description`, and the Oqtane theme `Name`. Do not change these to the project name.
- The **Access database and VBA macro add-ins keep `Explore the World` as their window title** (excluded from the project-name rule); their tests still assert `Explore the World`.
- The one-word **`ExploreTheWorld`** is reserved for code identity — namespaces, assembly names, the repo/folder names, `%LocalAppData%\JBC\ExploreTheWorld` paths, and the `JBC_ExploreTheWorld` database. Do not "fix" these.
- **Ribbon/tab labels stay short** ("ETW (VSTO)", "ETW (Web)", "ETW …") because ribbon space is tight. VSTO ribbon-tab UI tests assert the short "ETW …" names — keep them in sync.
- UI tests that assert on the window/pane title expect the **`ETW <ProjectName>` string** (WinForms/MAUI `MainForm_Tests`/`MainPage_Tests` assert e.g. `ETW WinFormApp`, `ETW MauiApp.WinUI`); update the Designer / `AppTitle` / `Title` string and the test together.

## Office Add-in Patterns

### Office.js Property Loads
- Verify every property/event name against the official Office.js reference (learn.microsoft.com/javascript/api) before using it in `load()`. An invalid name (e.g. `shapeType` — PowerPoint uses `type`) leaves the *entire* load unfilled, and the first later property read throws "The property 'x' is not available".
- Gate version-dependent properties with `Office.context.requirements.isSetSupported("{Host}Api", "1.x")` instead of probing with try/catch around `context.sync()` — a failed sync discards everything else queued in the same batch. Known gates: PowerPoint shape alt text + `pageSetup` = 1.10, `presentation.properties` = 1.7.
- PowerPoint slide size lives on `presentation.pageSetup.slideWidth/slideHeight` (1.10), *not* on `DocumentProperties`. Word `DocumentProperties` has no `wordCount`/`paragraphCount`/`pageCount`; derive counts from `body`. Word `revisionNumber` is a string.

### Office.js Event Registration
- Use named handler functions (not arrow functions) so `EventHandlerResult.remove()` can cleanly unregister them.
- Keep `Word.run()` / `Excel.run()` / `PowerPoint.run()` for event registration — handlers persist after the run completes.
- Guard against undefined event properties: `context.document[prop]` may be `undefined` if the runtime API version is below the requirement (or the event is preview-only). Log a console warning and skip gracefully.
- Pass an `eventKeys` string array from C# to JS so the caller controls which events are registered. The C# side exposes a `List<EventDescriptor>` with `Key`, `Label`, `Enabled` that the user can toggle before starting.
- Selection change in Word and PowerPoint uses the common API `Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, …)` — neither host has a production rich-API document/presentation `onSelectionChanged` event. `PowerPoint.Presentation.onSlideSelectionChanged` is PowerPointApi BETA (preview-only).
- Word content-control deleted/entered/exited events are per-`ContentControl` instance (WordApi 1.5), not on `Word.Document`; register them on each existing control.
- Excel worksheet events live on the worksheet collection (`onActivated` etc., ExcelApi 1.7); collection-level `onSelectionChanged` requires 1.9 — fall back to the active sheet's event (1.7).

## Do Not Do

- Do not reverse the layer dependency rule (AL → BL → DL → CL).
- Do not put HTTP/file/DB access in BL services — use DL repo interfaces.
- Do not skip the `._netF.sln` build after shared code changes.
- Do not use `I` prefix for interfaces — use `_Interface` suffix.
- Do not treat README.md as coding instructions.
- Do not duplicate code between `.NET 10` and `.NET Framework` projects — use shared-link compilation instead.
