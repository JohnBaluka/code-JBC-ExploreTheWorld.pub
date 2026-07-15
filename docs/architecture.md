# ExploreTheWorld Architecture Standards

## Application Purpose

The ExploreTheWorld applications demonstrate the various ways to program, integrate, and automate Microsoft Office applications. The same set of features is implemented across multiple modalities so that the code, patterns, and trade-offs of each approach can be compared side-by-side:

| Modality | Project(s) | Notes |
|---|---|---|
| **Access VBA** | `VBA/Access/ExploreTheWorld.accdb` | Desktop Access database application; common code shared with VBA Macro Add-ins |
| **Office VBA Macro Add-ins** | `VBA/{Word\|Excel\|PowerPoint}/ExploreTheWorld.{dotm\|xlsm\|pptm}` | Macro-enabled Office files; reuse code from the Access VBA modules |
| **WinForms App** | `AL.WinFormApp` (net10.0), `AL.WinFormApp._netF` (net481) | Desktop WinForms host; Blazor Hybrid content reused in Blazor Web Add-ins |
| **VSTO-Style COM Add-ins** | `AL.MsOffice{Host}VstoAddIn` (net10.0), `AL.MsOffice{Host}VstoAddIn._netF` (net481) | COM-hosted add-ins; net10.0 opens floating Blazor (WebView) forms; net481 uses traditional WinForms + Custom Task Panes |
| **Blazor Web Add-ins** | `AL.MsOffice{Host}BlazorWebAddIn` + `.Client` | Office.js task pane add-ins using Blazor WebAssembly |
| **Blazor Standalone (PWA)** | `AL.BlazorWebApp.ClientOnly` | WASM-only progressive web app; no server required at runtime |

### Code Sharing Hierarchy

```
Access VBA modules  ──────────────────►  VBA Macro Add-ins
                                          (reuse export/API modules)

AL.WinFormApp ExploreTheWorld_Form  ───►  VSTO COM Add-ins (net10.0)
(single BlazorWebView host; all pages     (floating Blazor forms only)
 via Radzen sidebar navigation)

AL.WinFormApp._netF Forms   ──────────►  VSTO COM Add-ins (net481)
(CountriesNowSpace_Form,                  (task pane + floating WinForms)
 Watcher_Forms, Watcher_UserControls)

AL.BlazorWebApp Blazor components  ────►  Blazor Web Add-ins
(CountriesNowSpace, Watcher)              (Office.js task pane pages)
```

## Application Feature Matrices

The matrices below record, per application, which options are actually implemented and selectable in the code (verified against the code on 2026-07-05). The number in each column header is a **sort key** used to order the options consistently across documents and UIs.

Cell values:
- **Yes** — implemented and selectable/functional in that app.
- **No** — the platform could support it, but it is not implemented (or not reachable from the UI).
- **No (NA)** — not applicable to that platform, or deliberately excluded by design (see footnotes).

### Applications × DB Provider Options

| AppName | 200 : ServerDB : InMemoryDb | 201 : ServerDB : AccessDb | 202 : ServerDB : SqliteDb | 203 : ServerDB : SqlServerDb | 210 : BrowserDB : LocalStorageDb | 211 : BrowserDB : IndexedDb | 212 : BrowserDB : SessionStorageDb |
|---|---|---|---|---|---|---|---|
| BlazorWebApp | Yes | Yes | Yes | Yes | No ¹ | No ¹ | No ¹ |
| BlazorWebApp.ClientOnly | Yes | No | No | No | Yes | Yes | Yes |
| Oqtane | No (NA) | No (NA) | No (NA) | Yes ² | No | No | No |
| VBA AddIn | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| Access Database | No (NA) | Yes ³ | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| WinFormApp | Yes | Yes | Yes | Yes | No (NA) | No (NA) | No (NA) |
| WinFormApp._netF | No (NA) ⁴ | Yes | No ⁵ | Yes | No (NA) | No (NA) | No (NA) |
| MauiApp.WinUI | Yes | Yes | Yes | Yes | No (NA) | No (NA) | No (NA) |
| MauiApp.Mac | Yes | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| MauiApp.Droid | Yes | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| MauiApp.iOS | Yes | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| VstoAddin | No ⁶ | No ⁶ | Yes ⁶ | No ⁶ | No (NA) | No (NA) | No (NA) |
| VstoAddin._netF | No (NA) ⁴ | No | Yes ⁷ | No | No (NA) | No (NA) | No (NA) |
| WebAddIn | No ⁸ | No (NA) | No (NA) | No (NA) | No ⁸ | No ⁸ | No ⁸ |
| ExportData | Yes ⁹ | Yes ⁹ | Yes ⁹ | Yes ⁹ | No (NA) | No (NA) | No (NA) |
| SaveAsJson | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |

¹ This row describes the server leg. The hybrid app's WASM client (`AL.BlazorWebApp.Client/Program.cs`) registers the three browser providers with the runtime switcher, so after the InteractiveAuto transition they are selectable in the browser.
² Shares the Oqtane host's own SQL Server LocalDB (`DefaultConnection`); fixed, no switcher.
³ Native Access `.accdb` tables (the VBA application's own database), not an EF Core provider.
⁴ No net481 InMemory `_Impl` project exists — the EF Core 3.x provider set is Sqlite/SqlServer/Access only.
⁵ `SqliteDb_Impl._netF` exists but is not wired — `Program.cs` offers only `AccessDb`/`SqlServerDb` via config (`DbProvider`), no runtime switcher.
⁶ Fixed provider hardcoded to `SqliteDb` in `Addin.cs`; no switcher (the `AccessDb` branch is unreachable dead code).
⁷ Fixed `SqliteDb` via the net481 static factory (`ExploreTheWorldSqliteDb.CreateFactory`); no switcher.
⁸ The web add-ins register no CountriesNowSpace DB provider at all — the CountriesNow pane reads the API directly.
⁹ All four registered via the keyed switcher; selected per run with `--provider` (no UI).

### Applications × Export Data Options

| AppName | 120 : Export Data : Direct | 121 : Export Data : COM | 122 : Export Data : Interop | 123 : Export Data : Dynamic | 124 : Export Data : NetOffice | 125 : Export Data : OpenXML | 126 : Export Data : OfficeJs |
|---|---|---|---|---|---|---|---|
| BlazorWebApp | No (NA) ¹ | No (NA) | No (NA) ¹ | No (NA) ¹ | No (NA) ¹ | Yes | No (NA) |
| BlazorWebApp.ClientOnly | No (NA) ¹ | No (NA) | No (NA) ¹ | No (NA) ¹ | No (NA) ¹ | Yes | No (NA) |
| Oqtane | No (NA) ¹ | No (NA) | No (NA) ¹ | No (NA) ¹ | No (NA) ¹ | Yes | No (NA) |
| VBA AddIn | No (NA) ² | No (NA) ² | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| Access Database | No ³ | Yes ³ | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| WinFormApp | No (NA) ⁴ | No (NA) | Yes | Yes | Yes | Yes | No (NA) |
| WinFormApp._netF | No (NA) ⁴ | No (NA) | Yes | Yes | Yes | Yes | No (NA) |
| MauiApp.WinUI | No (NA) ⁴ | No (NA) | Yes | Yes | Yes | Yes | No (NA) |
| MauiApp.Mac | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No ⁵ | No (NA) |
| MauiApp.Droid | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No ⁵ | No (NA) |
| MauiApp.iOS | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No ⁵ | No (NA) |
| VstoAddin | No (NA) ⁴ | No (NA) | Yes | Yes | Yes | Yes | No (NA) |
| VstoAddin._netF | No (NA) ⁴ | No (NA) | No (NA) ⁶ | No (NA) ⁶ | Yes ⁶ | No (NA) ⁶ | No (NA) |
| WebAddIn | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | Yes ⁷ |
| ExportData | No (NA) ⁴ | No (NA) | Yes | Yes | Yes | Yes | No (NA) |
| SaveAsJson | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |

¹ Browser-hosted: the shared export "Library" dropdown lists **only OpenXML** — it is populated from `OfficeExport_AppService__Interface.SupportedLibraries`, and `BrowserExport_AppService` returns just `["OpenXML"]` (desktop hosts' `OfficeExport_AppService` returns Interop/Dynamic/NetOffice/OpenXML). `BrowserExport_AppService` still coerces any other value to OpenXML defensively.
² The Office macro add-ins contain Save-As-JSON macros only; there is no country-data export to call via `Application.Run`.
³ The Access export (`ETW__cns_Ms{Host}_VBA`) always automates a `New {Host}.Application` COM object; `Application.Run` into the macro add-ins is used only by the Save-As-JSON watcher forms.
⁴ `Direct` is Save-As-JSON-only by design (`ExportMethod_Enum`: "Direct is Save-As-JSON only, so it is not an export method").
⁵ Export UI hidden (`Layout_AppService.ShowExportOptions = false`) and no `OfficeExport_AppService__Interface` registered on these heads.
⁶ Add-in export mode (`EnableAddinExportMode`) locks the library to NetOffice and exports into a new document in the running host application.
⁷ The CountriesNow pane's **Export to Document** button (on the WASM `.Client/Pages/CountriesNow` page, `@rendermode InteractiveWebAssembly` like the other Office.js pages) writes the loaded countries into the active document via Office.js (`{Host}CountriesExport__Interop` → `wwwroot/js/countries-export.js`): Word appends a Flag/Country/ISO2/ISO3 table, Excel writes the rows to the active worksheet at A1 (Flag/Country/ISO2/ISO3), and PowerPoint builds a title slide ("Explore the World" + count) followed by one centered slide per country (name, flag, ISO codes) — matching the other PowerPoint exports; when the deck is the default single starter slide it is reused as the title slide (no leading blank), otherwise a title slide is appended. When the **Flag images** toggle is on (default), each row/slide also gets its country's flag image (Wikimedia PNG thumbnail via `CL.FlagImageUrl_Helper`, fetched in-browser). Each `countries-export.js` reads the flag PNG's real pixel size from its IHDR header and sets the image width from that **aspect ratio** (the Office.js `lockAspectRatio`/height-only path does **not** rescale width), so flags are not distorted; the PowerPoint flag is a rectangle fill sized to the same aspect. Office.js is also used by the separate Save-As-JSON page.

### Applications × Save As JSON Options

| AppName | 310 : Save As JSON : Direct | 311 : Save As JSON : COM | 312 : Save As JSON : Interop | 313 : Save As JSON : Dynamic | 314 : Save As JSON : NetOffice | 315 : Save As JSON : OpenXML | 316 : Save As JSON : OfficeJs |
|---|---|---|---|---|---|---|---|
| BlazorWebApp | No (NA) ¹ | No (NA) | No (NA) ¹ | No (NA) ¹ | No (NA) ¹ | No (NA) ¹ | No (NA) |
| BlazorWebApp.ClientOnly | No (NA) ¹ | No (NA) | No (NA) ¹ | No (NA) ¹ | No (NA) ¹ | No (NA) ¹ | No (NA) |
| Oqtane | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| VBA AddIn | Yes (Macro) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| Access Database | Yes | Yes | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| WinFormApp | Yes | No (NA) | Yes | Yes | Yes | Yes | No (NA) |
| WinFormApp._netF | Yes | No (NA) | Yes | Yes | Yes | Yes | No (NA) |
| MauiApp.WinUI | Yes | No (NA) | Yes | Yes | Yes | Yes | No (NA) |
| MauiApp.Mac | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| MauiApp.Droid | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| MauiApp.iOS | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| VstoAddin | Yes | No (NA) | Yes | Yes | Yes | No (NA) ² | No (NA) |
| VstoAddin._netF | Yes | No (NA) | Yes | Yes | Yes | Yes ³ | No (NA) |
| WebAddIn | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | Yes |
| ExportData | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| SaveAsJson | Yes | No (NA) | Yes | Yes | Yes | Yes | No (NA) |

¹ The watcher pages render in these hosts, but there is no COM host to connect to — the Save button falls back to writing the watcher-state JSON, not a document.
² Hidden by design when `WatcherEvent_AppService.IsOfficeAddinHost` is true — OpenXML must close the active document, which an add-in host must keep open.
³ The `._netF` WinForms watcher form lists all five enum values and does **not** hide OpenXml in the add-in host (unlike net10) — a known inconsistency; OpenXml closes and reopens the active document.

### Applications × UI Implementation

| AppName | 110 : API UI : Form/Page | 111 : API UI : Pane | 251 : Slides UI : Form/Page | 252 : Slides UI : Pane | 301 : Watcher UI : Form/Page | 302 : Watcher UI : Pane |
|---|---|---|---|---|---|---|
| BlazorWebApp | Yes | No (NA) | Yes | No (NA) | No (NA) ¹ | No (NA) |
| BlazorWebApp.ClientOnly | Yes | No (NA) | Yes | No (NA) | No (NA) ¹ | No (NA) |
| Oqtane | Yes | No (NA) | Yes | No (NA) | No (NA) | No (NA) |
| VBA AddIn | No (NA) ² | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| Access Database | Yes | No (NA) | No (NA) | No (NA) | Yes | No (NA) |
| WinFormApp | Yes | No (NA) | Yes | No (NA) | Yes | No (NA) |
| WinFormApp._netF | Yes | No (NA) | No (NA) | No (NA) | Yes | No (NA) |
| MauiApp.WinUI | Yes | No (NA) | Yes | No (NA) | Yes | No (NA) |
| MauiApp.Mac | Yes | No (NA) | Yes | No (NA) | No (NA) | No (NA) |
| MauiApp.Droid | Yes ³ | No (NA) | Yes ³ | No (NA) | No (NA) | No (NA) |
| MauiApp.iOS | Yes | No (NA) | Yes | No (NA) | No (NA) | No (NA) |
| VstoAddin | Yes | No (NA) ⁴ | No (NA) | No (NA) | Yes | No (NA) ⁴ |
| VstoAddin._netF | Yes | Yes | No (NA) | No (NA) | Yes | Yes |
| WebAddIn | No (NA) | Yes | No (NA) | No (NA) | No (NA) | Yes ⁵ |
| ExportData | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |
| SaveAsJson | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) | No (NA) |

¹ The `/watcher-*` routes exist in the shared `AL.BlazorLib`, but the nav items are hidden (`Layout_AppService.ShowWatcherNavItems = false`) — there is no COM host in the browser.
² Ribbon buttons only (`RibbonUI.xml`); no forms.
³ Droid shares the same `Routes`/NavMenu/`Layout_AppService` configuration as Mac/iOS, so the Countries Now and Country Slides pages render there too.
⁴ net10 `comhost.dll` cannot activate UserControl subclasses — no Custom Task Panes; floating Blazor WebView forms only.
⁵ The Watcher pane in the web add-ins is the Events page (Office.js event log).

## Overview

ExploreTheWorld follows a **layered, service-oriented architecture** built with ASP.NET Core and Blazor technologies. The architecture is designed to support both .NET 10.0 (net10.0) and .NET Framework 4.8.1 (net481) through a shared-link compilation pattern, enabling code reuse across framework targets.

### Runtime hardening notes for the VSTO + Blazor path

The VSTO COM add-ins and Blazor Hybrid forms rely on the following runtime safeguards:

- The VSTO `ServiceProvider` registers `AddWindowsFormsBlazorWebView()`, `AddRadzenComponents()`, and `WatcherEvent_AppService()` so the floating Blazor (WebView) forms can instantiate their WebView-dependent components without crashing.

- **`AppContext.BaseDirectory` empty (net10.0 VSTO only):** `comhost.dll` does not initialize `APP_CONTEXT_BASE_DIRECTORY`. `BlazorWebView.StartWebViewCoreIfPossible()` calls `Path.GetRelativePath(AppContext.BaseDirectory, ...)` which throws `ArgumentException: The path is empty` when `BaseDirectory` is `""`, preventing the WebView from starting at all (form shows a blank window). Fixed at the very start of `OnConnection` with `AppContext.SetData("APP_CONTEXT_BASE_DIRECTORY", addinDir + Path.DirectorySeparatorChar)` where `addinDir = Path.GetDirectoryName(typeof(Addin).Assembly.Location)`.

- **`_framework/` files not in `wwwroot/` (net10.0 VSTO only):** `BlazorWebView.CreateFileProvider()` only creates a `PhysicalFileProvider(contentRootDir)` — it does **not** read `*.staticwebassets.runtime.json` at runtime and does **not** pick up `IFileProvider` from DI. `blazor.webview.js` and `blazor.modules.json` live in the NuGet package (`microsoft.aspnetcore.components.webview/{version}/staticwebassets/`) and are never in `wwwroot/` by default. Without them, `index.html` loads and shows "Loading..." forever (Blazor never initializes). Fix: the `CopyBlazorWebViewFrameworkFiles` MSBuild target (in the VSTO addin `.csproj`) copies them to `$(OutputPath)wwwroot\_framework\` after each build using `$(NuGetPackageRoot)microsoft.aspnetcore.components.webview\*\staticwebassets\*`.

Note: Task-pane registration (CLSID/ProgID, `CreateCTP`) and the ActiveX safety shim (`IObjectSafety`) are only required by the `._netF` COM add-ins (net481), which use `NetOfficeFw`'s `COMAddin` base class. The net10.0 add-ins open floating WebView forms only — there are no task panes. See repo memory `dotnet10-comhost-usercontrol-taskpane.md` for the full investigation.

## Core Architecture Layers

The application is organized into four primary layers, with specialized projects supporting each layer:

### 1. Common Layer (CL) - `JBC.ExploreTheWorld.CL`

**Purpose:** Shared utilities, extensions, and data models used across all layers

**Key Responsibilities:**
- Enum extensions and helper methods
- Tree node structures for hierarchical data
- Base row entities (audit logging support)
- Reusable utility classes
- No external dependencies

**Key Classes:**
- `TreeNode.cs` - Hierarchical tree structure with ParentID, ID, Text, Checked, Expanded, and Children properties
- `Enum_Extensions.cs` - Static methods for enum manipulation (GetEnumDisplayName, GetEnumDescription, etc.)
- `ILogger2.cs` - Custom logger interface for DI
- `_Row/_Row.cs` - Base entity class with GUID, Row_ID, Name; GetPrimaryKeyValue(), GetBusinessKeyValue() methods
- `RowLog_Row.cs` - Audit logging model for row changes
- `ColumnLog_Row.cs` - Audit logging for column-level changes

**Framework Support:** net10.0, net481

**Dependencies:** None (foundation layer)

---

### 2. Dependency Layer (DL) - `JBC.ExploreTheWorld.DL`

> **Naming note:** "DL" stands for **Dependency Layer**, not "Data Layer". DL projects encapsulate *external dependencies* the rest of the app must not bind to directly — databases, HTTP APIs, the file/OS system, Office document libraries, and **JavaScript interop**. Each dependency is wrapped behind an interface and injected via DI.

**Purpose:** Abstraction over external dependencies (data access, JS interop, system/network)

**Key Responsibilities:**
- Database context and repository patterns
- Entity models (inherit from CL._Row)
- Query abstractions
- CRUD operation templates

**Core Package References:**
- Microsoft.Extensions.Logging.Abstractions (10.0.9)
- System.Text.Json (10.0.9)

**Specialized DL Projects:**

#### `DL.MsJSInterop` and `DL.MsJSInterop.*` — JavaScript Interop

JS interop lives in dedicated dependency projects so `.razor` components never import JS modules or call `IJSRuntime` directly — they inject a typed `{Name}__Interop__Interface` registered in `Program.cs`.

- **`DL.MsJSInterop`** — *generic* browser interop. Hosts `wwwroot/js/download-file.js` and `layout.js` (ESM modules) plus their wrappers: `FileDownload__Interop` (`downloadFileFromBytes` / `downloadText`) and `Layout__Interop` (`getWindowWidth` / `watchWindowWidth`). Also provides `JsModuleInterop__Base`, the lazy ESM-import + cache-bust base class all interop services derive from.
- **`DL.MsJSInterop.RevealJs`** — wraps reveal.js. Contains `wwwroot/js/reveal-interop.js` (formerly `countrySlides.js`), the bundled `wwwroot/revealjs/**` library, and `RevealJs__Interop` (formerly `CountrySlides__Repo`). Consumed by `AL.BlazorLib`'s `CountrySlides__Component`.
- **`DL.MsOfficeApi.MsOfficeJs.{Word|Excel|PowerPoint}_Impl`** — per-host Office.js page interop (grouped under `DL.MsOfficeApi.MsOfficeJs.*` with the `_Impl` suffix per the Office-repo convention; formerly `DL.MsJSInterop.MsOffice{Word|Excel|PowerPoint}Js`). Each still references `DL.MsJSInterop` for the shared `JsModuleInterop__Base`. Each holds the page modules (`document-info.js`/`workbook-info.js`/`presentation-info.js`, `events.js`, `save-as-json.js`, `countries-export.js`, `home.js`) and one typed interop per module (e.g. `WordDocumentInfo__Interop`, `WordEvents__Interop`, `WordSaveAsJson__Interop`, `WordCountriesExport__Interop`), with the JSON result types (`{Host}DocumentInfo_Row`, `{Host}SaveAsJsonResult_Row`) shared by both sides. The Office.js-shaped rows and mappers used by Save-As-JSON live in the core `DL` project (`DL/MsOfficeApi/MsOfficeJs/{Host}/`, see below); `{Host}SaveAsJson__Interop` deserializes the collected data into those rows, maps them to the canonical `DL.MsOfficeApi.Ms{Host}` entities, and serializes with `MsOfficeJsonSerializer` so the web add-in output matches all other writers.

**Convention:** interop pair = `{Name}__Interop` (impl, derives `JsModuleInterop__Base`) + `{Name}__Interop__Interface`. Module assets are served from `_content/{AssemblyName}/js/`. These projects are net10/Blazor-only — there are no `._netF` variants. `SharedUtils.js` (`[JSImport]`/`JSHost`) and the Office `commands.js`/`*.lib.module.js` add-in infrastructure intentionally remain in the add-in projects.

#### `DL/MsOfficeApi/MsOffice`, `DL/MsOfficeApi/MsPowerPoint`, `DL/MsOfficeApi/MsExcel`, `DL/MsOfficeApi/MsWord` — canonical Save-As-JSON entities

The core `ExploreTheWorld.DL` project owns the **canonical "Save as JSON" schema**: one entity graph per Office host whose property names, types, and order follow the VBA object model ([learn.microsoft.com/office/vba/api](https://learn.microsoft.com/en-us/office/vba/api/overview/)). Every writer — VBA, NetOffice, OpenXML, and the Office.js web add-ins — produces this same schema so the `.json` files serialize and deserialize through the same classes.

- **Folders:** `_Entities/` (object graph; each entity derives its `_Fields` class), `_Fields/` (scalar VBA properties), `_Enums/` (`{Name}_Enum` with VBA member names). Every Ms enum includes `Undefined = -99`.
- **`MsOffice/MsOfficeJsonSerializer`** — the single serializer all .NET writers use: 2-space indent, CRLF, UTF-8 without BOM, trailing newline, default encoder, explicit nulls, `AllowTrailingCommas` when reading legacy files. Property order relies on System.Text.Json serializing the derived entity class (objects/lists) before the `_Fields` base class — the same order the VBA writers emit.
- **`MsOffice/MsOfficeUndefined`** — markers for properties a given API cannot provide: strings → `"**Undefined"`, enums/plain numerics → `-99`, booleans/dates/whole objects → `null`.
- **`MsOffice/MsOfficeJsonWriterOptions` + `BlobOutput_Enum` + `ImageBlob`** — image/blob output options shared by all writers. Default `Base64` embeds bytes in the JSON (`ImageBlob.Base64`); `SeparateFiles` writes files to `{jsonName}_Files/` beside the `.json` and stores the relative reference in `ImageBlob.FileName`.
- **`MsOfficeJs/{PowerPoint|Excel|Word}/`** — the Office.js-shaped rows (`{Host}…Js_Row`, matching the object graph collected by the web add-in `save-as-json.js` modules) and the `Ms{Host}JsMapper` classes that convert them to/from the canonical entities (missing Office.js properties get the Undefined markers).
- **Writers:** `DL.MsOfficeApi.NetOffice_Impl/JsonWriters/Ms{Host}JsonWriter` (live COM object via NetOffice; late-bound so it reads everything VBA can), `DL.MsOfficeApi.Interop_Impl/JsonWriters/Ms{Host}JsonWriter` (same live-COM read, strongly typed against the `Microsoft.Office.Interop` PIAs), `DL.MsOfficeApi.OpenXml_Impl/JsonWriters/Ms{Host}JsonWriter` (file on disk, no Office required; VBA-only properties get the Undefined markers), the VBA `MSO_Ms{Host}_JsonWriter.bas` modules, and the `DL/MsOfficeApi/MsOfficeJs/{Host}` mappers used by the web add-in interops.

#### `DL.MsSystem` - System-Level Data Access
- System utilities and platform-specific patterns
- OS-level file/registry operations
- System configuration access

#### `DL.MsSystemNet` - Network & HTTP Utilities
- HTTP client patterns
- System.Net.Http.Json for JSON serialization
- REST client abstractions

**Framework Support:** net10.0 (primary), net481 (via DL._netF)

**Layer Dependencies:** CL

---

### 3. Business Logic Layer (BL) - `JBC.ExploreTheWorld.BL`

**Purpose:** Business rules, validations, and domain logic

**Key Responsibilities:**
- Business rule implementation
- Data validation logic
- Service methods combining DL queries with business logic
- Manager classes orchestrating complex operations

**Core Package References:**
- Microsoft.Extensions.Logging.Abstractions (10.0.9)

**Folder Structure:**
- `/_Services` - business-logic orchestrator services (`{Name}__Service`); the leading-underscore folder is not part of the namespace
- Domain-specific folders for feature grouping

**DB-Backed Service Pattern (check-DB → API-fallback → persist → return with source):**

BL services act as orchestrators that check the local SQL database before calling an external API.
Each BL service follows this naming convention:
- Implementation: `{Name}__Service` (concrete — no interface of its own; the injected DL repo interfaces are the mocking seam)
- BL defines no interfaces of its own; the host-implemented export seam is a DL contract, `DL.MsOfficeApi.MsOfficeExportRepoFactory__Interface`

The return type `DataResult_Row<T>` (in `JBC.ExploreTheWorld.CL`) wraps the data list with a `DataSource_Enum` (`Api` or `Database`) so callers can display where the data came from.

```csharp
// CL - shared across all layers
public enum DataSource_Enum { Api, Database }

public class DataResult_Row<T>
{
    public List<T> Data { get; set; } = new();
    public DataSource_Enum Source { get; set; }
}

// BL - orchestrates DB check before API call (concrete service, no interface of its own)
public class CountriesNowSpaceManager__Service
{
    // Inject: CountriesNowSpaceApi_Interface (API) +
    //         CountriesNowSpaceApiManager__Repo__Interface (DL.CountriesNowSpaceData DB manager)

    public async Task<DataResult_Row<CountryBasic_Row>> GetAllCountriesAsync()
    {
        // 1. Check DB
        var dbRows = await _dbManager.GetCountriesAsync();
        if (dbRows.Count > 0)
            return new DataResult_Row<CountryBasic_Row> { Source = DataSource_Enum.Database, Data = /* map */ };

        // 2. Call API, persist, return
        var apiRows = await _apiService.GetAllCountriesAsync();
        foreach (var row in apiRows) await _dbManager.CreateCountryAsync(/* map */);
        return new DataResult_Row<CountryBasic_Row> { Source = DataSource_Enum.Api, Data = apiRows };
    }

    public Task ClearAllDataAsync() => _dbManager.ClearAllAsync();
}
```

DI registration order in `Program.cs`:
1. `services.AddExploreTheWorld{Provider}Db(...)` — registers `IDbContextFactory` + DL repo classes (provider sub-project)
2. `services.AddTransient<CountriesNowSpaceManager__Service>()`
3. `serviceProvider.EnsureExploreTheWorldDbCreated()` — calls `EnsureCreated()` before first use

**Usage Pattern:** BL services live in `BL/_Services/` and are concrete (no paired interface — the injected DL interfaces are the mocking seam). A leading-underscore folder is not part of the namespace, so the namespace is just `JBC.ExploreTheWorld.BL`.
```csharp
namespace JBC.ExploreTheWorld.BL
{
    public class FeatureManager__Service
    {
        private readonly FeatureApi_Interface _apiService;
        private readonly FeatureApiManager__Repo__Interface _dbManager;

        public FeatureManager__Service(FeatureApi_Interface apiService,
            FeatureApiManager__Repo__Interface dbManager)
        {
            _apiService = apiService;
            _dbManager  = dbManager;
        }
    }
}
```

**Framework Support:** net10.0 (primary), net481 (via BL._netF)

**Layer Dependencies:** CL, DL, DL.CountriesNowSpaceApi

**FlagImageManager — cached country flag images (check-cache → download → persist → return with source):**

`FlagImageManager__Service` (BL) resolves PNG flag images for ISO2 country codes with the same
cache-first shape as the DB-backed managers, but against a binary image store instead of a DB:

1. Check `FlagImageStore__Repo__Interface` (DL) — returns cached bytes → `Source = Database`.
2. On a miss, resolve the country's Wikimedia SVG flag URL via
   `CountriesNowSpaceManager__Service.GetCountryFlagsAsync()` (itself DB-backed against
   `cns_CountryFlag`), convert it to a rasterized PNG thumbnail URL with
   `CL.FlagImageUrl_Helper.GetPngThumbnailUrl(...)` (Wikimedia `/thumb/.../330px-*.svg.png` —
   Wikimedia only serves a fixed list of thumbnail widths, see https://w.wiki/GHai),
   download it via `FlagImageDownload__Repo__Interface` (DL), persist to the store, and return
   with `Source = Api`. The download repo sets a `User-Agent` header — Wikimedia rejects
   UA-less requests with HTTP 403 (browser hosts send their own UA instead).

Store implementations (registered per host):

| Host type | Implementation | Location |
|-----------|----------------|----------|
| Desktop/server (web app server, WinForms, MAUI, VSTO add-ins, Oqtane) | `FlagImageStore_FileSystem__Repo` (core DL) | `%LocalAppData%\JBC\ExploreTheWorld\FlagImages\{ISO2}.png` |
| Browser (Blazor WASM, ClientOnly PWA) | `FlagImageStore_Browser__Repo` (`DL.MsJSInterop/FlagImageCache`, wraps the `FlagImageCache__Interop` IndexedDB module) | IndexedDB `etw-flag-images` |

The Access VBA exports use the same file-cache folder via `ETW__FlagImages.bas`
(`URLDownloadToFile` + the same URL derivation), so all Windows hosts share one download.

Consumers:
- `MsOfficeExportManager__Service` accepts an optional `FlagImageManager__Service` and
  enriches `MsOfficeCountry_Row` rows (`FlagPng` bytes + `FlagFilePath`) before delegating to the
  Word/Excel/PowerPoint writers. OpenXML writers embed the bytes as image parts; NetOffice/VBA
  writers `AddPicture` from the cached file path (falling back to `%TEMP%\ETW_FlagImages`).
- `BrowserExport_AppService` performs the same enrichment for in-browser OpenXML exports.
- `CountrySlides__Component` loads flags in the background after reveal.js initializes and swaps
  the emoji flag for the real image as each one resolves (emoji remains the fallback).

---

### 4. Application Layer (AL) - `JBC.ExploreTheWorld.AL`

**Purpose:** Application orchestration and service composition

**Key Responsibilities:**
- Service registration and configuration
- Application initialization
- Dependency injection setup

**Core Package References:**
- Microsoft.Extensions.Logging.Abstractions (10.0.9)

**Folder Structure:**
- Minimal core layer; most functionality in specialized AL projects

**Framework Support:** net10.0 (primary), net481 (via AL._netF)

**Layer Dependencies:** CL, BL, DL

---

## Specialized Application Layer Projects

### AL.BlazorLib - Reusable Blazor Component Library

**Purpose:** Shared Blazor components, layouts, and infrastructure

**SDK Type:** Razor SDK (Razor Class Library)

**Key Responsibilities:**
- Custom component base classes
- Application layouts and shell components
- Service injection for Blazor context
- Global component imports

**Core Package References:**
- Radzen.Blazor (referenced from external workspace)
- Microsoft.AspNetCore.Components (10.0.9)
- Microsoft.AspNetCore.Components.Web (10.0.9)
- Microsoft.Extensions.Localization (10.0.9)

**Key Classes:**
- `_Shared/Base__RadzenComponent.cs` - Custom base class for Radzen components
  - Inherits RadzenComponent
  - Adds Hidden parameter and GetHiddenStyle() helper
  - Injects IJSRuntime, NavigationManager, and Radzen services
  - Provides centralized component configuration
- `_Services/RenderMode_AppService.cs` - Manages IComponentRenderMode
- `_Services/WatcherEvent_AppService.cs` - Singleton shared-state service; WinForms Watcher forms push Office events into it; Blazor Watcher pages subscribe via `event Action? {Host}StateChanged`. The Open Documents combobox is bidirectional: selecting a document in the UI calls `Word_SetActiveDocument`, and incoming `WordActivateDocumentAction` events update the UI to track which document Word has focused. `WatcherEventToggle.Log = true` by default so all events are logged from the start. `IsOfficeAddinHost` is set to `true` by the VSTO add-ins (see `Addin.cs`) so the watcher's "Save As JSON" UI hides the **OpenXML** method — OpenXML must close the active document, which an add-in host must keep open.

**Watcher & Save-As-JSON behavior (net10 WebView watchers):**
- There is no input-file textbox. The **Connect** button attaches to a running Office instance (or launches one). The Open Documents dropdown selects/activates a document.
- Event log lines carry parameter detail, e.g. `[Event] WindowSelectionChange: Aruba` (Word/PowerPoint selection text, Excel selected-range address).
- The default `.json` output name follows the standard: the active document's file name with the selected write method inserted before the extension and `.json` appended (e.g. `filename-NetOffice.pptx.json`), recomputed whenever the active document **or the selected method** changes (`SaveAsJson_Helper.BuildDefaultPath`). Cloud (AutoSave/OneDrive) documents report an `https://d.docs.live.net/...` (or `...-my.sharepoint.com`) path; `OneDriveLocalPath_Helper` maps it to the local OneDrive sync folder (via the `%OneDriveConsumer%`/`%OneDriveCommercial%`/`%OneDrive%` environment variables) so the default is a writable local path — when no local copy exists, the default falls back to the Documents folder with the document's file name.
- **Save-As-JSON** methods: `Direct` (calls the VBA add-in macro), `Interop` (raw COM object cast to the `Microsoft.Office.Interop` PIAs), `Dynamic` (late-bound `dynamic` COM), `NetOffice` (live COM object), `OpenXML` (reads the file on disk — the active document is **closed then reopened** around the read; cloud/AutoSave files with an `https` path are read from a **temporary local copy** via `SaveCopyAs` instead — Word has no `SaveCopyAs`, so cloud Word documents report a clear error). OpenXML is not offered in add-in hosts. Save duration is logged. The whole flow lives in `Ms{Host}_JsonWriter.WriteOpenXmlFromRunningApp` (`AL.WinFormsLib/_Watcher/`), shared by `ExploreTheWorld_Form`, the WebView watcher forms, and the `._netF` watcher forms.
- Log panels auto-scroll to the newest line (CSS `flex-direction: column-reverse`); the WinForms `ExportLog_Form` uses `RichTextBox.ScrollToCaret()`.
- `_Shared/Main_Layout.razor` - LayoutComponentBase with header placeholder and @Body
- `Countries/CountriesNow__Page.razor` - Countries/States grids + export section (type, library, path, Export, Clear Log, terminal log)
- `Watcher/Ms{Word|Excel|PowerPoint}_Watcher__Page.razor` - Full watcher UI (open files dropdown, save path, Save as JSON, Clear Log, terminal log) backed by `WatcherEvent_AppService`

**Folder Structure:**
```
AL.BlazorLib/
  _Imports.razor          # Global usings for all components
  _Shared/
    Base__RadzenComponent.cs
    Main_Layout.razor         # LayoutComponentBase: sidebar + header with renderModeBadge + breakpointBadge
    Main_Layout.razor.cs      # Code-behind: IJSRuntime for breakpoint, RendererInfo for render mode
  _Services/
    RenderMode_AppService.cs
    WatcherEvent_AppService.cs  # Singleton event bridge (WinForms → Blazor)
  Countries/
    CountriesNow__Page.razor      # Countries/States grids + export section
    CountriesNow__Page.razor.cs
  Watcher/
    MsWord_Watcher__Page.razor    # Full watcher UI backed by WatcherEvent_AppService
    MsWord_Watcher__Page.razor.cs
    MsExcel_Watcher__Page.razor
    MsExcel_Watcher__Page.razor.cs
    MsPowerPoint_Watcher__Page.razor
    MsPowerPoint_Watcher__Page.razor.cs
  Managers/               # Manager components for specific domains
  MainView/              # Main application view components
  wwwroot/
    css/                 # Stylesheets
    @* JS interop moved out of AL.BlazorLib:
       download-file.js + layout.js → DL.MsJSInterop (FileDownload__Interop / Layout__Interop)
       countrySlides.js + revealjs/ → DL.MsJSInterop.RevealJs (RevealJs__Interop) *@
```

**Framework Support:** net10.0 (primary)

**Layer Dependencies:** AL, BL, CL, DL and specialized DL projects

---

### AL.BlazorLib._radzen - Radzen-Focused Variant

**Purpose:** Radzen-specific component library for specialized UI needs

**SDK Type:** Razor SDK

**Extends:** AL.BlazorLib functionality

**Use Cases:**
- Advanced Radzen component patterns
- Radzen-specific layout templates
- Specialized dialogs and forms

**Framework Support:** net10.0

**Layer Dependencies:** Same as AL.BlazorLib

---

### AL.BlazorLib.Server._radzen - Server-Side Variant

**Purpose:** Server-side Blazor with Radzen components

**SDK Type:** Razor SDK

**Use Cases:**
- Server-side rendering scenarios
- Real-time applications
- Applications requiring immediate server access

**Framework Support:** net10.0

**Layer Dependencies:** AL.BlazorLib and dependencies

---

### AL.BlazorWebApp - Hybrid Web Application Server

**Purpose:** ASP.NET Core host for hybrid Blazor apps

**SDK Type:** Web SDK (ASP.NET Core 10.0)

**Responsibilities:**
- HTTP request handling and routing
- Static asset serving
- Hybrid interactive rendering: `InteractiveAuto` render mode (server-side on first load, transitions to WASM after download)
- Service registration for web context

**Program.cs Pattern:**
```csharp
var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddInteractiveWebAssemblyComponents();

builder.Services.AddRadzenComponents();

// Register application services
builder.Services.AddScoped<OfficeExport_AppService__Interface, BrowserExport_AppService>();
builder.Services.AddSingleton<WatcherEvent_AppService>();
builder.Services.AddScoped<RenderMode_AppService>();

var app = builder.Build();

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
        typeof(AL.BlazorWebApp.Client.App).Assembly,
        typeof(Routes).Assembly);

app.Run();
```

**Service registration notes:**
- `OfficeExport_AppService__Interface` must be registered on the server so static pre-render does not throw DI errors. The export action requires user interaction (interactivity) so the pre-render pass never calls it.
- `WatcherEvent_AppService` must be registered as singleton. It is idle in the web app context (no WinForms host), but Watcher pages in BlazorLib inject it.
- `typeof(Routes).Assembly` adds `AL.BlazorLib` so its `@page` components are discoverable. `typeof(Client.App).Assembly` adds the WASM client for interactive rendering.
- `app.MapStaticAssets()` is required in .NET 10 to serve WASM framework files alongside `UseStaticFiles()`.
- **Logging**: `appsettings.json` sets default log level; `appsettings.Development.json` adds `"DetailedErrors": true` for browser-visible Blazor error details.
- For `AL.BlazorWebApp.ClientOnly`, appsettings files go in `wwwroot/` (served to WASM): `wwwroot/appsettings.json` and `wwwroot/appsettings.Development.json`.

**Folder Structure:**
```
AL.BlazorWebApp/
  App.razor               # Root component and HTML layout
  Program.cs              # Server initialization
  Pages/                  # Server-side pages and components
  Components/             # Shared components
  wwwroot/               # Static assets
    css/
    js/
```

**Core Package References:**
- Microsoft.AspNetCore.Components.WebAssembly.Server (10.0.9)
- Microsoft.AspNetCore.HeaderPropagation (10.0.9)

**Framework Support:** net10.0

**Layer Dependencies:** AL.BlazorLib, AL.BlazorWebApp.Client, AL, BL, CL, DL.*

---

### AL.BlazorWebApp.Client - WebAssembly Client Application

**Purpose:** Client-side Blazor WebAssembly application

**SDK Type:** BlazorWebAssembly SDK

**Responsibilities:**
- Client-side routing and page components
- User interface rendering in the browser
- Client-side state management
- API communication with server

**Program.cs Pattern:**
```csharp
var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

// Register client-side services
builder.Services.AddScoped(sp => new HttpClient 
    { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddScoped<RenderMode_AppService>();

await builder.Build().RunAsync();
```

**Folder Structure:**
```
AL.BlazorWebApp.Client/
  App.razor               # Root component
  Program.cs              # Client initialization
  Pages/
    Index.razor           # Home page
    Home.razor
  _Imports.razor         # Global usings for client
  wwwroot/               # Static assets
    css/
    js/
```

**Core Package References:**
- Microsoft.AspNetCore.Components.WebAssembly (10.0.9)

**Framework Support:** net10.0 (primary)

**Layer Dependencies:** AL.BlazorLib, AL, BL, CL, DL.*

---

### AL.BlazorWebApp.ClientOnly — Standalone WebAssembly PWA

**Purpose:** Self-contained Blazor WebAssembly app (no ASP.NET server at runtime); distributable as a static-file site or PWA.

**SDK Type:** `Microsoft.NET.Sdk.BlazorWebAssembly`

**Key Differences from AL.BlazorWebApp.Client:**
- Has its own `index.html` (not served by a server project)
- Uses `WebAssemblyHostBuilder.CreateDefault` and mounts `Routes` (from `AL.BlazorLib`) directly
- Includes service worker (`service-worker.js` / `service-worker.published.js`) and `manifest.webmanifest` for PWA support
- References the same `AL.BlazorLib`, `BL`, `DL.*` stack as the hybrid client

**Program.cs Pattern:**
```csharp
var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<Routes>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

// Direct HTTP to countriesnow.space (no server proxy)
builder.Services.AddScoped<CountriesNowSpaceApi_Interface>(
    _ => new CountriesNowSpaceApi__Repo(new HttpClient
        { BaseAddress = new Uri("https://countriesnow.space/api/v0.1/") }));

// Null-object repo: SQLite not available in WASM; BL always falls through to the API
builder.Services.AddScoped<CountriesNowSpaceApiManager__Repo__Interface,
                            CountriesNowSpaceApiManager__WasmNoCache__Repo>();
builder.Services.AddScoped<CountriesNowSpaceManager__Service>();

// Browser OpenXML export (DocumentFormat.OpenXml in-memory + JS Interop download)
builder.Services.AddScoped<OfficeExport_AppService__Interface, BrowserExport_AppService>();
builder.Services.AddScoped<RenderMode_AppService>();

await builder.Build().RunAsync();
```

**Framework Support:** net10.0

**Layer Dependencies:** AL.BlazorLib, AL, BL, CL, DL.*

---

### AL.MauiLib + AL.MauiApp.* — .NET MAUI Blazor Hybrid (Multi-Project)

The MAUI app follows the **MAUI multi-project** layout (one shared library + one head project per platform), modeled on the VS "Multi-Project" MAUI template:

| Project | TFM | Role |
|---------|-----|------|
| `AL.MauiLib` | `net10.0` | Shared library (targets no single platform, hence `Lib` not `App` — mirrors the `AL.WinFormsLib` / `AL.WinFormApp` split): `App`, `MainPage` (BlazorWebView host), `MauiNewWindow_AppService`, and `MauiProgramExtensions.UseSharedMauiLib()` which registers the platform-neutral **BL orchestrators** + interops (Radzen, `CountriesNowSpaceManager__Service`, `FlagImageManager__Service`, JS interops, idle `WatcherEvent_AppService`). References **no DL repo project** — each head registers the concrete DL repos (`CountriesNowSpaceApi__Repo`, `FlagImageStore_FileSystem__Repo`, `FlagImageDownload__Repo`) and references `DL.CountriesNowSpaceApi` itself. Platform-neutral references only — no server EF providers / NetOffice / Dynamic. |
| `AL.MauiApp.WinUI` | `net10.0-windows10.0.19041.0` | Windows head. Full feature set: all four server EF providers switched at runtime via the BL `DbProviderSwitcher__Service`; Word/Excel/PowerPoint Watchers with all five Save-As-JSON methods (`_Services/MauiWatcher_AppService` via NetOffice COM, mirroring `ExploreTheWorld_Form`); Export API Data via the real Office export pipeline (`OfficeExport_AppService` shared-linked from `AL.WinFormsLib`; the `MsOfficeExportRepoFactory` + Save-As-JSON writers come from the `DL.MsOfficeApi_Impl` composition project this head references — Interop/Dynamic/NetOffice/OpenXML); `MauiNewWindow_AppService`. Loads `appsettings.json` from the exe directory (MAUI does not load it automatically). Unpackaged (`WindowsPackageType=None`) so FlaUI can launch the exe. `AnyCPU` builds pin `RuntimeIdentifier=win-x64` so plain `dotnet build` works. |
| `AL.MauiApp.Droid` | `net10.0-android` | Android head. InMemoryDb only; watcher nav + export UI hidden; `NullNewWindow_AppService`. |
| `AL.MauiApp.iOS` | `net10.0-ios` | iOS head. Same reduced registrations as Droid. |
| `AL.MauiApp.Mac` | `net10.0-maccatalyst` | Mac Catalyst head. Same reduced registrations as Droid but registers `MauiNewWindow_AppService` (Catalyst supports MAUI multi-window). |

Head-specific registrations (everything not in `UseSharedMauiLib()`): `DbProvider_AppService`, DB provider(s), `Layout_AppService` flags (`ShowWatcherNavItems`, `ShowExportOptions`), `OfficeExport_AppService__Interface`, and the `NewWindow_AppService__Interface` implementation. The non-Windows heads call `AddExploreTheWorldInMemoryDb()` (from `DL.CountriesNowSpaceData.InMemoryDb_Impl`), which binds `CountriesNowSpaceApiManager__Repo__Interface` directly — no switcher needed. Each head owns its `wwwroot/index.html` plus the platform bootstrap files (`MainActivity`/`MainApplication`, `AppDelegate`/`Main`, WinUI `App.xaml`); the shared project deliberately contains no `Platforms/` folder.

`Microsoft.AspNetCore.Components.WebView.Maui` is referenced by the shared project with `ExcludeAssets="build;buildTransitive"` (compile-time only) and by each head normally — referencing its build assets from both sides produces duplicate static-web-asset errors.

**Framework Support:** net10.0 (+ platform TFMs per head)

**Layer Dependencies:** AL.BlazorLib, AL, BL, CL, DL.* (WinUI head adds the four DL.CountriesNowSpaceData.*Db_Impl EF providers + the DL.MsOfficeApi_Impl composition project, which owns the DL.MsOfficeApi.*_Impl references)

---

### OfficeExport_AppService__Interface — OpenXML Browser Export

**Interface:** `OfficeExport_AppService__Interface` (in the core `ExploreTheWorld.AL` project — all AL
interfaces and data objects live there; the `AL.*` projects only implement them)

**Browser implementation:** `BrowserExport_AppService` (in `AL.BlazorLib`, file under `_Services/`)

> The desktop WinForms implementation `OfficeExport_AppService` (in `AL.WinFormsLib`) does **not**
> instantiate DL export repos directly. It delegates to the business-layer
> `MsOfficeExportManager__Service` (in `BL/_Services/`), which selects
> the target host and asks a host-supplied `MsOfficeExportRepoFactory__Interface` for the concrete
> DL repo. This keeps UI items free of DL references (UI → BL → DL).

The export service builds Office documents in memory using `DocumentFormat.OpenXml 3.2.0` (WASM-compatible, pure .NET) and triggers a browser download by injecting `FileDownload__Interop__Interface` (which wraps `downloadFileFromBytes` in `DL.MsJSInterop/wwwroot/js/download-file.js`).

| Export type | Document type | Content type |
|-------------|--------------|--------------|
| Word | DOCX with bordered table (Country, ISO2, ISO3) | `application/vnd.openxmlformats-officedocument.wordprocessingml.document` |
| Excel | XLSX with single "Countries" sheet | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| PowerPoint | PPTX with title slide + up to 10 data slides (20 countries/slide) | `application/vnd.openxmlformats-officedocument.presentationml.presentation` |

**Registration:** `builder.Services.AddScoped<OfficeExport_AppService__Interface, BrowserExport_AppService>()` in `AL.BlazorWebApp/Program.cs` (server), `AL.BlazorWebApp.Client/Program.cs`, and `AL.BlazorWebApp.ClientOnly/Program.cs`.

**Host-capability flags on `OfficeExport_AppService__Interface`** — the shared `CountriesNow__Component` export UI is driven by two properties so it adapts to the registered service (no per-host UI code):
- `IReadOnlyList<string> SupportedLibraries` — populates the "Library" picklist. `BrowserExport_AppService` returns `["OpenXML"]` (the only method in the WASM sandbox); the desktop `OfficeExport_AppService` returns `Interop/Dynamic/NetOffice/OpenXML`.
- `bool SavesToFilePath` — `false` for `BrowserExport_AppService` (no local file system: the field is labeled **"File Name"** and defaults to a bare file name; the document streams to the browser as a **download**), `true` for `OfficeExport_AppService` (field labeled **"File Path"**, defaults to the file in the user's Documents folder).

So the four browser hosts (`AL.BlazorWebApp`, `AL.BlazorWebApp.ClientOnly`, Oqtane, `AL.BlazorLib._radzen`) show OpenXML-only + a download file name; WinForms/MAUI show all four libraries + a local save path.

**Desktop export document layout (DL.MsOfficeApi.*_Impl writers).** The desktop "Export Data" writers (OpenXML / Interop / Dynamic / NetOffice, selected per the WinForms/MAUI Export Library combo) all produce the **same visual output** for a given host so the four approaches compare like-for-like:

- **Word** — a title paragraph ("Countries Now", no export timestamp) followed by a bordered 4-column table: **Flag · Country · ISO2 · ISO3**. Flag images are inserted at a fixed height (0.25″ / 18 pt) with the **aspect ratio locked** so the width scales proportionally.
- **Excel** — a single "Countries" sheet with a 4-column header (**Flag · Country · ISO2 · ISO3**); column A holds the flag image (aspect locked, ~16 pt tall) per data row.
- **PowerPoint** — a title slide ("Explore the World" + country count, **no timestamp**) followed by **one slide per country**, each with the country name (large, centered), a large centered flag image, and the ISO2 / ISO3 codes — mirroring the `/country-slides` reveal.js page. Positions are computed from the runtime slide size (COM writers) or the declared 9,144,000 × 5,143,500 EMU slide (OpenXML). The COM writers create the presentation **with a window** (`Presentations.Add(msoTrue)`); a windowless presentation makes PowerPoint fail SaveAs/Close/Quit with COM errors.

Flag bytes/paths come from `MsOfficeExportManager__Service`'s `FlagImageManager` enrichment (see the FlagImageManager section): OpenXML embeds the `FlagPng` bytes; the COM writers `AddPicture` from the cached file path. The separate in-browser `BrowserExport_AppService` (table above) keeps its own simpler layout.

The **Access VBA** country exports (`VBA/Access/…/ETW__cns_Ms{Word|Excel|PowerPoint}_VBA.cls`, which automate a new Office application from Access) and the **Office.js web add-in** exports (footnote ⁷ under the Export Data matrix) follow the **same conventions**: Word/Excel produce a Flag/Country/ISO2/ISO3 table with aspect-locked flags; PowerPoint produces a title slide ("Explore the World" + count, no timestamp) plus one centered slide per country (name, large flag, ISO codes). VBA flags come from the shared `%LocalAppData%` file cache via `ETW__FlagImages.bas`.

**Namespace alias pattern (required to avoid CS0104 ambiguity):**
```csharp
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
```

---

### CountriesNowSpaceApiManager__WasmNoCache__Repo — WASM Null-Object DB Repo

**Location:** `DL/_Repos/CountriesNowSpaceApiManager__WasmNoCache__Repo.cs`

**Purpose:** Satisfies `CountriesNowSpaceApiManager__Repo__Interface` in WASM without EF Core SQLite (which requires native libraries unavailable in WebAssembly). All read methods return empty collections; write methods are no-ops. This forces the BL manager to always call the live API.

**When to use:** Register this in any WASM project (`AL.BlazorWebApp.Client`, `AL.BlazorWebApp.ClientOnly`) instead of `CountriesNowSpaceApiManager__Repo`.

---

## Dependency Flow

```
┌─────────────────────────────────────────┐
│         AL (Application Layer)          │
├─────────────────────────────────────────┤
│ AL.BlazorWebApp  AL.BlazorLib  (servers)│
├─────────────────────────────────────────┤
│              BL (Business Logic)        │
├─────────────────────────────────────────┤
│           DL (Dependency Layer)         │
│  DL.MsJSInterop[.RevealJs]              │
│  DL.MsOfficeApi[.*_Impl]                │
│   DL.MsSystem  DL.MsSystemNet           │
├─────────────────────────────────────────┤
│           CL (Common Layer)             │
└─────────────────────────────────────────┘
```

**Rules:**
- CL depends on nothing
- DL depends only on CL
- BL depends on DL and CL
- AL depends on AL, BL, DL, and CL
- Lower layers never reference upper layers
- Services injected via DI to enforce abstraction

---

## Design Principles

### 1. Single Responsibility
Each layer and project has a single, well-defined responsibility:
- CL: Shared models and utilities only
- DL: Persistence and data access
- BL: Business rules and validation
- AL: Application orchestration

### 2. Dependency Injection
All services registered in container at startup:
```csharp
builder.Services.AddScoped<YourService__Service__Interface, YourService>();
```

### 3. Interface-Based Design
- Define interfaces in higher-level projects
- Implement in lower-level projects
- Inject via interface types

### 4. Async/Await Throughout
- Database operations async
- Network calls async
- Component lifecycle async

### 5. Separation of Concerns
- Components focused on UI
- Managers focused on business rules
- Repositories focused on data access
- Services focused on cross-cutting concerns

---

## Service Injection Pattern

### BL Manager Registration
```csharp
builder.Services.AddScoped<JBC.ExploreTheWorld.BL.ManagerName.FeatureManager>();
```

### DL Repository Registration
```csharp
builder.Services.AddScoped<JBC.ExploreTheWorld.DL.RepositoryName.IFeatureRepository, 
                           JBC.ExploreTheWorld.DL.RepositoryName.FeatureRepository>();
```

### Logging Registration
```csharp
builder.Services.AddLogging();  // Auto-configured by AddRazorComponents()
```

### Usage in Components
```csharp
@inject BL.ManagerName.FeatureManager FeatureManager

@code {
    protected override async Task OnInitializedAsync()
    {
        var result = await FeatureManager.GetDataAsync();
    }
}
```

---

## Framework Support Strategy

### .NET 10.0 (net10.0) - Primary
- All projects target net10.0 by default
- Uses latest C# features and APIs
- Blazor WebAssembly primary rendering platform

### .NET Framework 4.8.1 (net481) - Legacy Support
- Implemented via shared-link compilation
- Source files linked from net10.0 projects (not duplicated)
- Projects: CL._netF, DL._netF, BL._netF, AL._netF
- Compatibility packages: System.Net.Http (4.3.4)

**Shared-Link Pattern:**
```xml
<ItemGroup>
    <Compile Include="..\ProjectName\**\*.cs" />
</ItemGroup>
<ItemGroup>
    <Compile Remove="obj/**" />
    <Compile Remove="bin/**" />
</ItemGroup>
```

This allows 100% code reuse between frameworks without maintaining separate files.

---

## Build Targets

Four solution files provided for different build scenarios:

1. **JBC.ExploreTheWorld.sln** - Complete stack (all projects including test projects under a Tests solution folder)
   - All core layers, all specialized projects, all Office add-ins, all test projects
   - Recommended for full application development

2. **JBC.ExploreTheWorld.AL.BlazorLib._radzen.sln** - Radzen focus
   - **Apps:** `AL.BlazorLib._radzen`, `AL.BlazorLib.Server._radzen`
   - **Libs:** `AL`, `AL.BlazorLib`, `BL`, `CL`, `DL`, `DL.CountriesNowSpaceApi`,
     `DL.CountriesNowSpaceData` + all `*Db_Impl` providers, `DL.MsJSInterop`,
     `DL.MsJSInterop.RevealJs`, `DL.MsOfficeApi.OpenXml_Impl`, `DL.MsSystem`, `DL.MsSystemNet`
   - **Tests:** `IntegrationTests`, `OpenXmlLibTests`, `RazorTests`, `UnitTests`
   - Lets the `.razor` files be edited inside Radzen Blazor Studio (which cannot open Razor
     Class Libraries). The two projects form a **runnable hosted Blazor Web App** modeled on
     `AL.BlazorWebApp.Client` + `AL.BlazorWebApp`:
     - `AL.BlazorLib._radzen` — `Microsoft.NET.Sdk.BlazorWebAssembly` app that holds a
       **synchronized copy** of every `.razor` / `.razor.cs` / `.razor.css` and supporting
       `_Services`/`_Shared`/`Countries`/`Watcher` file from `AL.BlazorLib`. It references the
       same projects as `AL.BlazorLib` (plus the browser DB providers) and uses `RootNamespace`
       `JBC.ExploreTheWorld.AL.BlazorLib` (no `._radzen` suffix) so the copied components share
       their code-behind namespace; `AssemblyName` keeps the `._radzen` suffix. It does **not**
       reference `AL.BlazorLib` (standalone copy — no duplicate-type ambiguity).
     - `AL.BlazorLib.Server._radzen` — `Microsoft.NET.Sdk.Web` host (`App.razor` +
       `Program.cs`, `InteractiveAuto` render mode) that serves the WASM companion, so Radzen
       Studio can open and run it.
     - **Keep the copies in sync** whenever `AL.BlazorLib` `.razor`/code-behind changes.

3. **JBC.ExploreTheWorld.AL.BlazorWebApp.sln** - Web app focus
   - **Apps:** `AL.BlazorWebApp`, `AL.BlazorWebApp.ClientOnly`, `Oqtane.Server`
   - **Libs:** the Oqtane solution folder + projects; `AL`, `AL.BlazorLib`,
     `AL.BlazorWebApp.Client`, `BL`, `CL`, `DL`, `DL.CountriesNowSpaceApi`,
     `DL.CountriesNowSpaceData` + all `*Db_Impl` providers, `DL.MsJSInterop`,
     `DL.MsJSInterop.RevealJs`, `DL.MsOfficeApi.OpenXml_Impl`, `DL.MsSystem`, `DL.MsSystemNet`
   - **Tests:** `IntegrationTests`, `OpenXmlLibTests`, `OqtaneTests`, `RazorTests`,
     `UnitTests`, `WebAppTests`
   - For Blazor web + Oqtane application development (no WinForms/MAUI/VSTO/Office add-in projects)

4. **JBC.ExploreTheWorld._netF.sln** - .NET Framework
   - Only ._netF framework variants
   - For legacy framework support development

---

## External REST API Dependency Layer Projects

### Pattern: `DL.{Name}Api` Implementation Projects

For integrating external REST APIs, the **interface and data-row types live in the core
`ExploreTheWorld.DL` project**; the `DL.{Name}Api` project contains only the `HttpClient`
implementation. This follows the rule that *all DL interfaces and shared data objects belong in
`ExploreTheWorld.DL`, and the `DL.*` projects exist to implement those interfaces.*

```
DL/                                 # core Dependency Layer
  {Name}Api/
    _Interfaces/
      {Name}Api_Interface.cs        # Interface defining all API operations
    _Rows/
      {Entity}_Row.cs               # Response model classes with [JsonPropertyName] attributes

DL.{Name}Api/                       # implementation project (references DL)
  ExploreTheWorld.DL.{Name}Api.csproj
  {Name}Api__Repo.cs                # HttpClient-based implementation
```

The interface/row types keep the `JBC.ExploreTheWorld.DL.{Name}Api` namespace even though they
physically live in the `DL` project, so consumers' `using` statements are unaffected. Because
`DL.{Name}Api` references `DL` and SDK project references are transitive, every consumer of the
repo also resolves the interface/rows automatically.

**Key Design Decisions:**
- No `<ImplicitUsings>enable</ImplicitUsings>` — explicit `using` statements for ._netF shared-link compatibility
- Interface naming: `{Name}Api_Interface` (no `I` prefix — consistent with codebase convention)
- Repo naming: `{Name}Api__Repo` (double underscore — consistent with codebase convention)
- Row model naming: `{Entity}_Row` for all JSON response models
- Parameterless constructor creates its own `HttpClient` (sets `BaseAddress`); second constructor accepts injected `HttpClient`
- `[JsonPropertyName]` attribute on every property for explicit JSON mapping

**csproj Template:**
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <Nullable>enable</Nullable>
    <AssemblyName>JBC.ExploreTheWorld.DL.{Name}Api</AssemblyName>
    <RootNamespace>JBC.ExploreTheWorld.DL.{Name}Api</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="System.Net.Http.Json" Version="10.0.5" />
  </ItemGroup>
</Project>
```

**Framework Support:** net10.0 (primary), net481 (via `DL.{Name}Api._netF` shared-link project)

**._netF Compatibility Packages:**
- `System.Net.Http` 4.3.4
- `System.Net.Http.Json` 8.0.1 (last version supporting netstandard2.0)
- `System.Text.Json` 8.0.5

### Existing External API Projects

#### `DL.CountriesNowSpaceApi` — countriesnow.space
- Base URL: `https://countriesnow.space/api/v0.1/`
- GET for list endpoints; POST with `{ country }` JSON body for country-specific queries
- Generic response wrapper: `CountriesNowResponse_Row<T>` (Error, Msg, Data)
- Operations: `GetAllCountriesAsync()`, `GetCountryCapitalAsync(string)`, `GetCountryPopulationAsync(string)`, `GetCountryStatesAsync(string)`, `GetCountryFlagsAsync()`

---

## DL.MsOfficeApi.OpenXml_Impl / NetOffice_Impl / Dynamic_Impl / Direct_Impl — Office Document Repositories

These four DL projects each implement the three interfaces (`MsWord__Repo__Interface`, `MsExcel__Repo__Interface`, `MsPowerPoint__Repo__Interface`) defined in `DL/MsOfficeApi/`. Each interface exposes:

```csharp
Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log);
Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log);
```

`Dynamic_Impl` (formerly `Interop_Impl`) drives Office through late-bound `dynamic` COM. `Direct_Impl` is **JSON-write only**: `WriteDocumentJsonAsync` opens the document in its Office host and runs the ExploreTheWorld VBA writer macros via `Application.Run` (`DL.MsOfficeApi.Direct_Impl.MsOfficeDirectVbaRunner`, no NetOffice dependency), while `ExportAsync` throws `NotSupportedException` — the Direct method cannot build documents from country data.

The shared parameter type `MsOfficeCountry_Row(string Country, string Iso2, string Iso3, byte[]? FlagPng = null, string? FlagFilePath = null)` is a record defined in `DL/MsOfficeApi/_Rows/MsOfficeCountry_Row.cs` to avoid a circular DL ↔ DL.CountriesNowSpaceApi dependency. `FlagPng`/`FlagFilePath` carry the country's cached flag image (populated by the flag-image enrichment in `MsOfficeExportManager__Service` / `BrowserExport_AppService`); writers embed the flag in a leading **Flag** column (Word table cell / Excel one-cell-anchored drawing / PowerPoint per-row picture) and silently skip rows without an image. OpenXML writers consume the bytes; NetOffice writers `AddPicture` from `FlagFilePath` (or materialize the bytes to `%TEMP%\ETW_FlagImages` via `NetOfficeFlagImage_Helper`). Images preserve the source aspect ratio (`CL.PngImage_Helper` reads the PNG header; COM writers use `LockAspectRatio`).

| Project | Framework | Library | Notes |
|---------|-----------|---------|-------|
| `DL.MsOfficeApi.OpenXml_Impl` | net10.0 | `DocumentFormat.OpenXml 3.5.1` | Cross-platform; creates DOCX/XLSX/PPTX in-process |
| `DL.MsOfficeApi.OpenXml_Impl._netF` | net481 | `DocumentFormat.OpenXml 3.5.1` | Shared-link from `DL.MsOfficeApi.OpenXml_Impl` |
| `DL.MsOfficeApi.NetOffice_Impl` | net10.0-windows | NetOffice__10 project refs | COM automation; WriteDocumentJson reads open document object |
| `DL.MsOfficeApi.NetOffice_Impl._netF` | net481 | `NetOfficeFw.*` NuGet 1.9.10 | Shared-link from `DL.MsOfficeApi.NetOffice_Impl` |
| `DL.MsOfficeApi.Dynamic_Impl` | net10.0-windows | COM late-binding via `dynamic` | `Type.GetTypeFromProgID` + `Activator.CreateInstance` |
| `DL.MsOfficeApi.Dynamic_Impl._netF` | net481 | `Microsoft.CSharp` 4.7.0 | Shared-link from `DL.MsOfficeApi.Dynamic_Impl` |
| `DL.MsOfficeApi.Direct_Impl` | net10.0-windows | `dynamic` COM + VBA `Application.Run` | Runs `MSO_Ms{Host}.WriteActive…ToJsonFile`; JSON-write only, no NetOffice |
| `DL.MsOfficeApi.Direct_Impl._netF` | net481 | `Microsoft.CSharp` 4.7.0 | Shared-link from `DL.MsOfficeApi.Direct_Impl` |

**Save-As-JSON writers** — `DL.MsOfficeApi.{NetOffice|Interop|Dynamic}_Impl/JsonWriters/` build the same canonical `DL.Ms{Host}` entity graphs from a running Office document:

- `DL.MsOfficeApi.NetOffice_Impl/JsonWriters/` is **strongly typed** against the NetOffice object model (`Write{Document|Workbook|Presentation}ToJsonFile(NetOffice.{Host}Api.X, …)`). The handful of Office members the NetOffice wrappers do not expose (e.g. PowerPoint `Shape.GraphicStyle`/`InkXML`, Word `Shape.HasTextFrame`) are read late-bound from the wrapper's `UnderlyingObject`, keeping the JSON output identical to the VBA writers.
- `DL.MsOfficeApi.Interop_Impl/JsonWriters/` is **strongly typed** against the `Microsoft.Office.Interop` PIAs (`Write{Document|Workbook|Presentation}ToJsonFile(Microsoft.Office.Interop.{Host}.X, …)`) — the same object model as NetOffice with the same output. Members absent from the Office15 PIA (`Shape.Decorative`/`IsNarration`, `Presentation.AutoSaveOn`/`ReadOnlyRecommended`) are read late-bound via `((dynamic)x).Member`.
- `DL.MsOfficeApi.Dynamic_Impl/JsonWriters/` is **late-bound** (`dynamic`): `Write{Document|Workbook|Presentation}ComToJsonFile(object comObject, …)` accepts any Word/Excel/PowerPoint COM object (Interop PIA or late-bound). The Dynamic repos use it to implement `WriteDocumentJsonAsync` (visible app instance, open file, write, close).
- `DL.MsOfficeApi_Impl/_JsonWriters/Ms{Host}_JsonWriter` (the composition project — **not** the UI library) exposes `WriteNetOffice`/`WriteInterop`/`WriteDynamic`/`WriteOpenXml…`; `WriteInterop` unwraps the live NetOffice document to its raw COM object and casts it to the `Microsoft.Office.Interop` interface. `MsOfficeSaveAsJsonWriter` dispatches the watcher's method dropdown to the matching wrapper; the UI watcher forms call it through the host-set `MsOfficeSaveAsJsonWriterProvider` so `AL.WinFormsLib`/`._netF` reference no `_Impl` project. `MsOfficeJsonWriter_Helper` (also in the composition project, called directly by the VSTO add-in ribbon handlers) routes by object type: NetOffice objects go to the typed NetOfficeLib writers, anything else to the DynamicLib writers.
- Both share the `ComTryGet` helper semantics (a property read that throws serializes as `null`, invariant-culture formatting), matching `MSO_Ms{Host}_JsonWriter.bas`.

**Gold-standard alignment** — the VBA "Direct" writer output is the reference the .NET writers converge to (verified against `slides/HowToThinkInBlazor-*.pptx.json` exports; `src/Tests/WinFormAppTests/JsonWriters/` pins the behavior):

- `Single` values format with `"G7"` (VBA `CStr(Single)`: max 7 significant digits, e.g. `"50.61386"`, `"-2.147484E+09"`); dates written through `CStr` use `"M/d/yyyy h:mm:ss tt"`.
- A property read that throws serializes as `null` — the writers never substitute `0`/`""` defaults. The `DL.MsOfficeApi.MsPowerPoint` entity value fields are nullable (`int?`/`long?`/`bool?`) for this reason.
- Object getters are probed the way VBA evaluates them at the call site: `AnimationSettings.PlaySettings` (throws for non-media shapes) and `SoundEffect.Name` yield a `null` object, `Hyperlinks` with `Count == 0` serialize as `null`.
- `Shape.Vertices` follows `CStr` semantics: `""` for non-freeform shapes (COM returns Empty), `null` for freeform shapes (`CStr` of an array raises).
- Presentation `HasHandoutMaster`/`HasNotesMaster`/`HasVBProject` store the MsoTriState value (`msoTrue` = -1) even though the bindings surface them as `bool`.
- The NetOffice/Dynamic outputs for the same file are identical except `Name`/`FullName` when the file is opened Untitled (PowerPoint assigns a session-scoped `PresentationN` name).

**Direct instantiation pattern** — repos have no constructor dependencies; AL projects create instances with `new MsWord_OpenXml__Repo()`, `new MsWord_NetOffice__Repo()`, etc. and select the concrete type at runtime via `(ExportType, ExportMethod)` switch expressions in `ExportLog_Form.RunExportAsync`.

**DL.MsOfficeApi.NetOffice_Impl** also exposes a non-interface method `WriteJsonFromOpenDocument(Document doc, …)` for the Watcher JSON writer forms that hold a live NetOffice document object.

**Running-app export (Office add-in hosts)** — the three NetOffice repos additionally implement `MsOfficeRunningAppExport__Repo__Interface` (in `DL/MsOfficeApi/`):

```csharp
Task ExportToRunningAppAsync(object hostApplication, IList<MsOfficeCountry_Row> countries, Action<string> log);
```

This creates a **new document in an already-running host Office application** (the add-in's own instance) and leaves it open and unsaved — no file is written and the app is not launched or quit. `hostApplication` is typed as `object` so the interface and the BL manager stay NetOffice-free; the repo casts it to the concrete `NetOffice.{Host}Api.Application`. Because the passed COM object is bound to the caller's apartment, the method runs **synchronously on the caller's STA/UI thread** (no `Task.Run`). Only the NetOffice repos implement this interface — `MsOfficeExportManager__Service.ExportToRunningAppAsync` resolves the NetOffice repo and calls it; OpenXML/Dynamic/Direct are not applicable. The document-fill logic is shared with `ExportAsync` via a private `Fill{Document|Workbook|Presentation}` helper in each repo.

**Export visibility & timing** — the `ExportAsync` repos (NetOffice and Dynamic) keep the Office application **visible** (`Visible = true`) while exporting so the user can watch the document being built; only the internal JSON-writer path (which opens a throwaway hidden instance from a file) stays hidden. `ExportLog_Form` and the Blazor `CountriesNow__Component` wrap the export in a `Stopwatch` and log the elapsed seconds. The default export file name is `ETW_CountriesNow.{docx|xlsx|pptx}` (no timestamp).

**DI registration** (for apps that prefer it):
```csharp
services.AddOpenXmlLib();   // DL.MsOfficeApi.OpenXml_Impl.ServiceCollectionExtensions
services.AddNetOfficeLib();  // DL.MsOfficeApi.NetOffice_Impl.ServiceCollectionExtensions
services.AddDynamicLib();    // DL.MsOfficeApi.Dynamic_Impl.ServiceCollectionExtensions
services.AddDirectLib();     // DL.MsOfficeApi.Direct_Impl.ServiceCollectionExtensions
```

**BrowserExport_AppService** in `AL.BlazorLib/_Services/` delegates to the injected `MsOfficeDocument_Memory__Repo__Interface` (a core-DL contract implemented by `MsOfficeDocument_Memory__Repo` in `DL.MsOfficeApi.OpenXml_Impl`, registered by each browser host) which builds documents into a `MemoryStream` (byte[]) for browser downloads via JS interop. Because it takes the interface rather than newing the concrete repo, `AL.BlazorLib` references **no** DL `_Impl` project. `BrowserExport_AppService` maps `CountryBasic_Row` → `MsOfficeCountry_Row` before calling the repo so that `DL.MsOfficeApi.OpenXml_Impl` has no dependency on `DL.CountriesNowSpaceApi`.

---

## DL.CountriesNowSpaceData — EF Core Database Context and Provider Projects

`DL.CountriesNowSpaceData` holds the provider-agnostic `ExploreTheWorldDbContext` and the `EnsureExploreTheWorldDbCreated` helper. **It does not reference any EF Core provider package** — provider-specific packages live exclusively in the `_Impl` sub-projects below. The core `DL` project also defines `DbProviderNames` (at the `DL` project root, namespace `JBC.ExploreTheWorld.DL`) — the canonical provider-key string constants (`SqliteDb`, `LocalStorageDb`, …) shared by the `_Impl` registrations, the BL `DbProviderSwitcher__Service`, and every host's `AvailableProviders` list.

### Base projects

| Project | Framework | Package |
|---------|-----------|---------|
| `DL.CountriesNowSpaceData` | net10.0 | `Microsoft.EntityFrameworkCore` 10.0.9 |
| `DL.CountriesNowSpaceData._netF` | net481 | `Microsoft.EntityFrameworkCore` 3.1.32 |

`DL.CountriesNowSpaceData._netF` uses the shared-link pattern to compile all `.cs` files from `DL.CountriesNowSpaceData`, excluding `ServiceCollectionExtensions.cs` and the EF Core 5+/7+-incompatible manager repo. It provides:
- `ExploreTheWorldDbContextFactory` — takes `DbContextOptions<T>` directly (constructed by each provider's static `CreateFactory` method).
- `ServiceCollectionExtensions.EnsureExploreTheWorldDbCreated(factory)` — synchronous helper that calls `EnsureCreated()` on a newly created context.

### Provider sub-projects (net10.0)

Each `_Impl` project owns exactly one `CountriesNowSpaceApiManager__Repo__Interface` implementation and exposes two registration styles on its `ServiceCollectionExtensions`:
- **Single-provider** `AddExploreTheWorld{Provider}Db(...)` — binds the interface directly (`AddDbContextFactory` + `AddTransient<Repo__Interface, Repo>`). Used by hosts with one fixed provider (Office add-ins, MAUI mobile/Mac, Oqtane).
- **Keyed** `AddExploreTheWorld{Provider}DbProvider(...)` — registers the repo as a keyed service under its `DbProviderNames` key, for hosts that switch at runtime via the BL `DbProviderSwitcher__Service` (see below).

| Project | Package | Single-provider method |
|---------|---------|--------|
| `DL.CountriesNowSpaceData.SqliteDb_Impl` | `Microsoft.EntityFrameworkCore.Sqlite` 10.0.9 | `AddExploreTheWorldSqliteDb(connStr)` |
| `DL.CountriesNowSpaceData.SqlServerDb_Impl` | `Microsoft.EntityFrameworkCore.SqlServer` 10.0.9 | `AddExploreTheWorldSqlServerDb(connStr)` |
| `DL.CountriesNowSpaceData.AccessDb_Impl` | `EntityFrameworkCore.Jet` 10.0.0 | `AddExploreTheWorldAccessDb(dbPath)` |
| `DL.CountriesNowSpaceData.LocalStorageDb_Impl` | `Blazored.LocalStorage` 4.0.1 | `AddExploreTheWorldLocalStorageDb()` |
| `DL.CountriesNowSpaceData.SessionStorageDb_Impl` | `Blazored.SessionStorage` 2.4.0 | `AddExploreTheWorldSessionStorageDb()` |
| `DL.CountriesNowSpaceData.IndexedDb_Impl` | `Microsoft.JSInterop` 10.0.9 | `AddExploreTheWorldIndexedDb()` |
| `DL.CountriesNowSpaceData.InMemoryDb_Impl` | `Microsoft.EntityFrameworkCore.InMemory` 10.0.9 | `AddExploreTheWorldInMemoryDb(dbName?)` |

Each EF Core `_Impl` (Sqlite / SqlServer / Access / InMemory) additionally owns its typed `ExploreTheWorldDbContext__{Provider}__Factory` (single factory → at the `_Impl` project root, no `_Factories` folder), its thin `CountriesNowSpaceApiManager__{Provider}__Repo` subclass, and a `TryEnsureExploreTheWorld{Provider}DbCreated(IServiceProvider)` startup helper. There are **no aggregator projects** — the former `ServerDb` / `BrowserStorageDb` switcher projects were replaced by the single generic BL `DbProviderSwitcher__Service`.

**AccessDb** requires Microsoft Access Database Engine (ACE OLEDB 12.0 or 16.0) installed and `System.Data.OleDb` (>= 9.0.0) referenced — `System.Data.OleDb` was removed from the BCL in .NET 5 and must be added explicitly as a NuGet package. Default path: `%LocalAppData%\JBC.ExploreTheWorld\etw.accdb` (created by EFCore.Jet on first use; separate from the VBA application database).

**LocalStorageDb** references `DL` (entity types + repo interface) but **not** `DL.CountriesNowSpaceData` — it contains no EF Core DbContext. Uses `ILocalStorageService` to persist entities as JSON lists keyed by `etw.cns.*` keys. Registers as `AddScoped` (not `AddSingleton`) since `IJSRuntime` is scoped.

**SessionStorageDb** is identical to LocalStorageDb in structure but uses `ISessionStorageService` (Blazored.SessionStorage). Data is scoped to the browser tab session and cleared when the tab is closed.

**IndexedDb** uses a custom JS interop repo (`IndexedDb_Repo`, in `IndexedDb_Impl/_Repos/`) backed by a minimal ES module (`wwwroot/js/indexedDb.js`) that wraps the browser IndexedDB API as a key-value store. Serialization to/from JSON strings is done in C#. Data persists across page reloads and is scoped to the browser origin. The JS module is served as a static web asset at `_content/JBC.ExploreTheWorld.DL.CountriesNowSpaceData.IndexedDb_Impl/js/indexedDb.js`.

**InMemoryDb** requires no external database. Data is stored in-process and lost when the process exits. Suitable for development, demos, and unit/integration tests. The optional `databaseName` parameter (default `"JBC.ExploreTheWorld"`) scopes the in-memory store — pass a distinct name per test class to isolate test runs.

### Provider sub-projects (net481)

`._netF` projects use the static-factory pattern (no DI `AddDbContextFactory`):

| Project | Package | Static class |
|---------|---------|--------------|
| `DL.CountriesNowSpaceData.SqliteDb_Impl._netF` | `Microsoft.EntityFrameworkCore.Sqlite` 3.1.32 | `ExploreTheWorldSqliteDb.CreateFactory(connStr)` |
| `DL.CountriesNowSpaceData.SqlServerDb_Impl._netF` | `Microsoft.EntityFrameworkCore.SqlServer` 3.1.32 | `ExploreTheWorldSqlServerDb.CreateFactory(connStr)` |
| `DL.CountriesNowSpaceData.AccessDb_Impl._netF` | `EntityFrameworkCore.Jet` 3.1.1 | `ExploreTheWorldAccessDb.CreateFactory(dbPath)` |

### DB Provider Selection

Provider is selected at startup via `appsettings.json` (or `App.config` for net481). The initial value sets which provider is active on first launch; the user can switch at runtime via the header dropdown without restarting the app. **`InMemoryDb` is the default on every host where it is available** (the switchable net10 hosts) — it needs no external database, so the app runs out of the box; the config value or `?? DbProviderNames.InMemoryDb` code fallback can override it.

```json
{
  "DbProvider": "InMemoryDb",
  "ConnectionStrings": {
    "SqliteDb": "",
    "SqlServerDb": "Server=.;Database=JBC_ExploreTheWorld;Trusted_Connection=True;TrustServerCertificate=True;",
    "AccessDb": ""
  }
}
```

Empty-string values fall back to built-in defaults: SQLite → `%LocalAppData%\JBC.ExploreTheWorld\etw.db`; Access → `%LocalAppData%\JBC.ExploreTheWorld\etw.accdb`; SQL Server → `Server=.;…;TrustServerCertificate=True;`.

### Runtime provider switching (BL `DbProviderSwitcher__Service`)

`AL.BlazorWebApp`, `AL.WinFormApp`, `AL.MauiApp.WinUI`, and `AL.ExportData.ConsoleApp` register each EF Core provider as a **keyed** implementation, then register the BL switcher as the `CountriesNowSpaceApiManager__Repo__Interface`. `DbProviderSwitcher__Service` (in `BL/_Services/`) resolves the keyed implementation matching the active provider name at call time and delegates every repo call to it.

The switcher reads the active name via an injected `Func<string>` delegate, so neither DL nor BL references the AL `DbProvider_AppService` directly:

```csharp
var dbProviderName = config["DbProvider"] ?? DbProviderNames.InMemoryDb;
var dbProvider_AppService = new DbProvider_AppService
{
    ProviderName       = dbProviderName,
    AvailableProviders = [DbProviderNames.InMemoryDb, DbProviderNames.AccessDb, DbProviderNames.SqliteDb, DbProviderNames.SqlServerDb],
};
services.AddSingleton(dbProvider_AppService);

// One keyed registration per provider (each _Impl owns its keyed repo + typed factory).
services.AddExploreTheWorldSqliteDbProvider(config.GetConnectionString("SqliteDb"));
services.AddExploreTheWorldSqlServerDbProvider(config.GetConnectionString("SqlServerDb"));
services.AddExploreTheWorldInMemoryDbProvider();
services.AddExploreTheWorldAccessDbProvider(config.GetConnectionString("AccessDb"));

// Generic switcher — resolves the active keyed repo per call (in ExploreTheWorld.BL).
services.AddCountriesNowSpaceDbSwitcher(() => dbProvider_AppService.ProviderName);
```

After the DI container is built, call each registered provider's `TryEnsure…` helper once. Each runs `EnsureCreated()` for that provider; unavailable providers (SQL Server not installed, Access engine missing, etc.) are silently skipped — the user will see an error only when they actually select one.

```csharp
serviceProvider.TryEnsureExploreTheWorldSqliteDbCreated();
serviceProvider.TryEnsureExploreTheWorldSqlServerDbCreated();
serviceProvider.TryEnsureExploreTheWorldInMemoryDbCreated();
serviceProvider.TryEnsureExploreTheWorldAccessDbCreated();
```

**Keyed-DI mechanism** — each `AddExploreTheWorld{Provider}DbProvider()` registers a typed `ExploreTheWorldDbContext__{Provider}__Factory` singleton (EF providers) plus `AddKeyedScoped<CountriesNowSpaceApiManager__Repo__Interface>(DbProviderNames.{Provider}, …)`. `DbProviderSwitcher__Service` injects `IServiceProvider` and resolves `GetRequiredKeyedService<…__Repo__Interface>(getActiveProviderName())` on each call. A host only registers the providers it supports (server hosts add the four EF providers; WASM hosts add the three browser providers), and the same switcher class works for both. Because the switcher implements a DL repo interface, it lives in BL as a composite/orchestrator (the same pattern as `MsOfficeExportManager__Service`), not in a DL project.

**Net481 apps** still use the single-provider static-factory pattern (no keyed registration, no switcher) and do not support runtime switching.

### Provider mapping per AL project

| AL Project | Providers |
|------------|-----------|
| `AL.BlazorWebApp` | SqliteDb / SqlServerDb / AccessDb / InMemoryDb (all switchable at runtime via the BL switcher) |
| `AL.WinFormApp` | SqliteDb / SqlServerDb / AccessDb / InMemoryDb (all switchable at runtime via the BL switcher) |
| `AL.MauiApp.WinUI` | SqliteDb / SqlServerDb / AccessDb / InMemoryDb (all switchable at runtime via the BL switcher) |
| `AL.MauiApp.Droid`, `AL.MauiApp.iOS`, `AL.MauiApp.Mac` | InMemoryDb (via `AddExploreTheWorldInMemoryDb()`, no switcher) |
| `AL.MsOffice{Host}VstoAddIn` (net10.0) | SqliteDb (fixed, hardcoded in `Addin.cs` — no switcher; the AccessDb branch is unreachable) |
| `AL.BlazorWebApp.Client`, `AL.BlazorWebApp.ClientOnly` | LocalStorageDb / SessionStorageDb / IndexedDb (switchable at runtime via the BL switcher) |
| `AL.WinFormApp._netF` | SqlServerDb / AccessDb |
| `AL.MsOfficeWordVstoAddIn._netF` | SqliteDb (static factory) |

### Browser storage runtime switching

`AL.BlazorWebApp.Client` and `AL.BlazorWebApp.ClientOnly` use the same BL switcher as the server hosts — they just register the three browser providers instead of the four EF providers:

```csharp
builder.Services.AddExploreTheWorldLocalStorageDbProvider();
builder.Services.AddExploreTheWorldSessionStorageDbProvider();
builder.Services.AddExploreTheWorldIndexedDbProvider();

var dbProvider_AppService = new DbProvider_AppService
{
    ProviderName       = DbProviderNames.InMemoryDb,
    AvailableProviders = [DbProviderNames.InMemoryDb, DbProviderNames.LocalStorageDb, DbProviderNames.IndexedDb, DbProviderNames.SessionStorageDb],
};
builder.Services.AddSingleton(dbProvider_AppService);

builder.Services.AddCountriesNowSpaceDbSwitcher(() => dbProvider_AppService.ProviderName);
```

The switcher reads `DbProvider_AppService.ProviderName` at call time via the `Func<string>` delegate, so switching the header dropdown takes effect immediately without restarting the app or re-creating the DI scope.

### DbProvider_AppService and provider dropdown

`AL/_Services/DbProvider_AppService.cs` (in the core `ExploreTheWorld.AL` project — it is a
framework-neutral state/data object, so it lives in AL core rather than `AL.BlazorLib`) is a singleton with:
- `ProviderName` — the active provider key (e.g. `"SqliteDb"`, `"LocalStorageDb"`)
- `AvailableProviders` — list of providers the user can switch between
- `OnProviderChanged` event — raised by `SetProvider(string)` for layout re-render
- `SetProvider(string)` — updates `ProviderName` and raises the event (no-op if unchanged)

`Main_Layout.razor` shows a `RadzenDropDown` (`id="dbProviderSelect"`) in the header. It is always enabled — all app types populate `AvailableProviders` at startup so the user can switch providers at runtime.

---

### AL.WinFormApp — BlazorWebView WinForms Host (net10.0, WebView only)

**Purpose:** Desktop WinForms host for the Blazor UI library

**SDK Type:** `Microsoft.NET.Sdk.Razor`

**Key Characteristics:**
- Hosts `AL.BlazorLib.Routes` via `BlazorWebView` control (docked Fill)
- `Program.cs` registers DI services including API repos: `AddTransient<{Name}Api_Interface, {Name}Api__Repo>()`, Radzen components (`AddRadzenComponents()`), and `AddSingleton<WatcherEvent_AppService>()`
- `wwwroot/index.html` with `<div id="app">` and `blazor.webview.js`
- `Main_Form` exposes two controls: **Countries Now (Blazor)** button and **Watcher (Blazor)** dropdown (Word / Excel / PowerPoint) — both open floating WebView forms
- Traditional WinForms (DataGridView) forms live in `AL.WinFormApp._netF` only

**WebView Forms (net10.0, in `_Forms/`):**
- `CountriesNowSpace_WebView_Form` — BlazorWebView host for CountriesNow__Page
- `MsWord_Watcher_WebView_Form`, `MsExcel_Watcher_WebView_Form`, `MsPowerPoint_Watcher_WebView_Form` — BlazorWebView hosts for Watcher pages

**Helper/Utility files (stay in AL.WinFormApp, shared-linked to _netF and VstoAddIn projects):**
- `_Forms/_Export/`: `ExportMenuHelper.cs`, `ExportMethod_Enum.cs`, `ExportType_Enum.cs`, `JsonWriteMethod_Enum.cs`
- `_Forms/_Watcher/`: `WatcherComHelper.cs`, `MsOfficeEvent_Record.cs`, `MsOfficeEvents_Repo.cs`, `MsOfficeJsonWriter_Helper.cs`, `MsWord_JsonWriter.cs`, `MsExcel_JsonWriter.cs`, `MsPowerPoint_JsonWriter.cs`

**NetOffice Integration:**

NetOffice is available on **both** targets:
- **net481 (`AL.WinFormApp._netF`):** Via `NetOfficeFw.Core/Excel/Word/PowerPoint` NuGet packages
- **net10.0 (`AL.WinFormApp`):** Via `NetOffice__10` project references (`code-zgh-NetOfficeFw__NetOffice__10\Source\*`)

`WatcherComHelper.GetActiveCom(progId)` abstracts platform differences for `GetActiveObject`:
- `NETFRAMEWORK`: `Marshal.GetActiveObject(progId)`
- non-Framework: P/Invoke `oleaut32.dll!GetActiveObject` via `Guid`

**NetOffice__10 Project References (in `JBC.ExploreTheWorld.sln`):**
- `NetOffice.Analyzers` — Roslyn analyzer for NetOffice event consistency
- `NetOffice` — Core runtime and COM proxy infrastructure
- `OfficeApi` — Shared Office COM types
- `VBIDEApi` — Visual Basic IDE COM types
- `ExcelApi`, `WordApi`, `PowerPointApi` — Application-specific COM wrappers

### AL.WinFormApp._netF — Traditional WinForms (No Blazor, net481)

**Purpose:** .NET Framework 4.8.1 WinForms app using traditional DataGridView controls (no BlazorWebView)

**SDK Type:** `Microsoft.NET.Sdk`

**Key Characteristics:**
- No DI framework — services instantiated directly in `Program.cs`
- `TabControl` with one tab per API source
- `DataGridView` for data display
- `async void` event handlers for `Load`/button click
- Row selection in one grid triggers async load in another (e.g., select country → load states)

**Owns all traditional Form/UserControl files (native, not linked):**
- `_Forms/CountriesNowSpace_Form.cs/.Designer.cs` — traditional DataGridView Countries form
- `_Forms/CountriesNowSpace_UserControl.cs/.Designer.cs` — UserControl variant (for net481 VSTO task panes)
- `_Forms/_Export/ExportLog_Form.cs/.Designer.cs/.resx`
- `_Forms/_Watcher/Ms{Word|Excel|PowerPoint}_Watcher_Form.cs/.Designer.cs` — standalone WinForms windows demonstrating NetOffice event watching + the JSON write methods (NetOffice / OpenXml / Dynamic / Direct)
- `_Forms/_Watcher/Ms{Word|Excel|PowerPoint}_Watcher_UserControl.cs` — UserControl variant shared to net481 VSTO add-ins

**Shared-link helper files (from AL.WinFormApp via `<Compile Include>` links):**
- `_Export/`: `ExportMenuHelper.cs`, `ExportMethod_Enum.cs`, `ExportType_Enum.cs`, `JsonWriteMethod_Enum.cs`
- `_Watcher/`: `WatcherComHelper.cs`, `MsOfficeEvent_Record.cs`, `MsOfficeEvents_Repo.cs`, `MsOfficeJsonWriter_Helper.cs`, `MsWord_JsonWriter.cs`, `MsExcel_JsonWriter.cs`, `MsPowerPoint_JsonWriter.cs`

---

## Access Database & VCS

The Access database lives at `VBA/Access/ExploreTheWorld.accdb`. Its structure and code is maintained via the **MSAccess-VCS Add-in** (v4.0.34), which exports all table data, form layouts, and VBA modules into the `VBA/Access/ExploreTheWorld.accdb.src/` source folder.

- **Local VCS tool:** `C:\Dev\github\joyfullservice\msaccess-vcs-addin`
- **GitHub:** https://github.com/joyfullservice/msaccess-vcs-addin
- **Rule:** When modifying the Access database, edit only the files in the `.accdb.src/` folder — never the `.accdb` binary directly. The VCS Add-in handles synchronising the source folder and the `.accdb` file.

### Access VBA Module Conventions

See also: repo memory `access-form-bas-standards.md` for `.bas`/`.cls` layout rules.

| Sub-folder | Contents |
|---|---|
| `tables/` | JSON table schema & data exports |
| `forms/` | Form layout `.bas` + VBA code-behind `.cls` |
| `modules/` | Standard `.bas` modules and class `.cls` modules |
| `queries/` | Saved query definitions |
| `relations/` | Relationship definitions |

### Navigation

Navigation uses the built-in Access **Custom Navigation Pane** (configured via Navigation Options). Forms are organized into custom groups in the Navigation Pane — no VBA code or AutoExec macro is required.

The former `Main` navigation form, `APP__Ribbon.bas` custom ribbon module, and AutoExec macro have all been removed.

**Module naming pattern:**
- `ETW__{API}_{OutputType}_{Implementation}` — export modules (e.g. `ETW__cns_MsWord_VBA`)
- `ETW__{API}_API` — load/clear modules (e.g. `ETW__cns_API`)
- APIs: `cns` = CountriesNow.space
- Output Types: `MsWord`, `MsExcel`, `MsPowerPoint`
- Implementations: `VBA` (Office automation, early-bound)

**VBA export class modules** (`VB_PredeclaredId = True`) allow default-instance call syntax:
```vba
ETW__cns_MsWord_VBA.Export Me.txt_Log, Me.Filter, Me.FilterOn
```

---

## VBA Macro Add-ins

For each Office host, a macro-enabled file holds a self-contained VBA add-in. Its primary purpose is to **export the active document's entire object model to a structured JSON file**. It also hosts CountriesNow and Watcher UserForms.

| Office App | File Extension | Location |
|---|---|---|
| Word | `.dotm` | `VBA/Word/ExploreTheWorld.dotm` |
| Excel | `.xlsm` | `VBA/Excel/ExploreTheWorld.xlsm` |
| PowerPoint | `.pptm` | `VBA/PowerPoint/ExploreTheWorld.pptm` |

All three macro add-ins are implemented. The canonical source for any shared module logic is the `.pptm` copy.

**Rule:** Like the Access database, edit only the `.src/` folder — never the macro file directly. The VBA source is stored in:
```
VBA/{Office App}/ExploreTheWorld.{ext}.src/
  *.bas     — standard modules
  *.cls     — class modules / host object (ThisDocument, ThisWorkbook, ThisPresentation)
  *.frm     — UserForms (layout + code)
  *.frx     — UserForm binary resources
```

### Module naming convention

| Module | Description |
|--------|-------------|
| `MSO_MsPowerPoint` | Entry-point wrapper. `WriteActivePresentation()` — prompts for output path via `InputBox` with datetime default, then calls the writer. `WriteActivePresentationToJsonFile(sOutputFilePath As String)` — path-only shim for `Application.Run` from C# or Access Direct mode. |
| `MSO_MsWord` | Same pattern: `WriteActiveDocument()` + `WriteActiveDocumentToJsonFile(sOutputFilePath As String)` |
| `MSO_MsExcel` | Same pattern: `WriteActiveWorkbook()` + `WriteActiveWorkbookToJsonFile(sOutputFilePath As String)` |
| `MSO_MsPowerPoint_JsonWriter` | Full-depth PPT object model → JSON serializer. Public entry: `WritePresentationToJsonFile(oPresentation As PowerPoint.Presentation, sOutputFilePath As String, Optional eBlobOutput As JsonBlobOutput = jsonBlobBase64, Optional sBlobFolderPath As String = "")` |
| `MSO_MsWord_JsonWriter` | Full-depth Word object model → JSON serializer. Public entry: `WriteDocumentToJsonFile(oDocument As Word.Document, sOutputFilePath As String, Optional eBlobOutput ..., Optional sBlobFolderPath ...)` |
| `MSO_MsExcel_JsonWriter` | Full-depth Excel object model → JSON serializer. Public entry: `WriteWorkbookToJsonFile(oWorkbook As Excel.Workbook, sOutputFilePath As String, Optional eBlobOutput ..., Optional sBlobFolderPath ...)` |
| `MSO_JsonWriterCore` | Shared JSON emitter used by all three writers: strictly valid JSON (a pending-line buffer decides commas), 2-space indent, CRLF, UTF-8 without BOM (ADODB.Stream), plus the `JsonBlobOutput` enum and base64/blob-folder helpers |

The writers emit the canonical `JBC.ExploreTheWorld.DL.Ms{Host}` entity schema — byte-identical to `MsOfficeJsonSerializer` output for the same object graph. String escaping uses `json_Encode()` from `JsonConverter.bas` (VBA-JSON v2.3.1, customized to match the System.Text.Json default encoder). Picture shapes are exported through the blob options: base64 embedded in the JSON (default) or separate files in `{jsonName}_Files/` that the JSON references.

### Ribbon: "ETW (VBA)" tab

Each macro add-in exposes a custom ribbon via `IRibbonExtensibility` with a single **Export** group:

| Control | ID | Description |
|---------|----|-------------|
| Tab | `tabETW` | "ETW (VBA)" |
| Group | `grpJsonExport` | "Export" |
| Button | `btnSaveAsJson` | "Save as JSON (Direct)" (`imageMso="FileSaveAs"`) — calls `SaveAsJson_OnClick` (in `RibbonHandlers.bas`) → `MSO_Ms{Host}.WriteActive{Document|Workbook|Presentation}()` |

The macro add-ins expose **only the Save-as-JSON feature** (the VBA "Direct" writer). CountriesNow and Watcher forms are not part of the VBA Macro Add-ins.

### Application.Run pattern

Because VBA's `Application.Run` can only pass simple types (not COM objects), the path-only shims on `MSO_MsX` are the correct entry points for remote invocation:
```vba
' From Access VBA or C# Direct mode
m_oPPT.Run "ExploreTheWorld.pptm!MSO_MsPowerPoint.WriteActivePresentationToJsonFile", sPath
m_oWord.Run "ExploreTheWorld.dotm!MSO_MsWord.WriteActiveDocumentToJsonFile", sPath
m_oExcel.Run "ExploreTheWorld.xlsm!MSO_MsExcel.WriteActiveWorkbookToJsonFile", sPath
```
The shims use `ActivePresentation` / `ActiveDocument` / `ActiveWorkbook` internally.

### Access DB copy

The Access DB maintains its own copies of all three writer modules (`MSO_MsPowerPoint_JsonWriter`, `MSO_MsWord_JsonWriter`, `MSO_MsExcel_JsonWriter`) plus the shared `MSO_JsonWriterCore` in `VBA/Access/ExploreTheWorld.accdb.src/modules/`. Each copy is identical to the corresponding macro add-in version except:
- Has `Option Compare Database` after the `Attribute VB_Name` line
- Is kept in sync with the corresponding add-in module manually

The `ETW__MsX_JsonWriter` modules have been removed — the `MSO_MsX_JsonWriter` pattern is the single authoritative approach across all modalities.

Each macro add-in source folder contains:

| File | Description |
|------|-------------|
| `This{Document|Workbook|Presentation}.cls` | Host class implementing `IRibbonExtensibility`; returns the embedded `RibbonUI.xml` |
| `RibbonUI.xml` + `RibbonHandlers.bas` | Ribbon XML (single **Export** group) + the `onAction` handler `SaveAsJson_OnClick` |
| `MSO_Ms{Host}.bas` | Entry-point wrapper (`WriteActiveX()` + `WriteActiveXToJsonFile(path)` shim) |
| `MSO_Ms{Host}_JsonWriter.bas` | Full-depth object model → JSON serializer (canonical `DL.Ms{Host}` schema) |
| `MSO_JsonWriterCore.bas` | Shared strict-JSON emitter + blob helpers (identical across add-ins and Access) |

---

## CountriesNowSpace Form Design

The CountriesNowSpace form is the primary data entry and export UI. Its design is consistent across all modalities with modality-specific adaptations:

### Common Controls (all modalities)
- **Load** button — calls the countriesnow.space API and populates the data grid / display
- **Clear** button — clears the local data
- **Export Type** ComboBox — selects the Office output format: Word / Excel / PowerPoint
- **Export** button — exports the loaded data to the selected format
- **Clear Log** button — clears the log text area
- **Log** text area — displays progress and error messages

### WinForms (AL.WinFormApp, AL.WinFormApp._netF) additional controls
- **Export Library** ComboBox — selects the .NET export library, ordered Interop / Dynamic / NetOffice / OpenXML (Interop uses the `Microsoft.Office.Interop` PIAs; Direct is Save-As-JSON only and is not an export library)
- **Export File** TextBox + **Browse** button — shows the output file path. It is **pre-filled on load** with the full default path in the user's **Documents** folder (`{MyDocuments}\ETW_CountriesNow-{library}.{ext}`, via `MsOfficeExportName_Helper.BuildFileName`) and re-computed when the Export Type or Library changes — unless the user has typed or Browsed to their own path, which is preserved. Browse opens in the current path's folder (or Documents). The path TextBox is anchored Left+Right so it **stretches** as the form widens, while Browse / Export / Clear Log stay anchored to the **right** edge.
- Log output styled as a terminal window (black background, green monospace text; `Consolas 9pt, BackColor=Black, ForeColor=Lime`)

### VSTO Add-in (VstoAddIn._netF, net481) adaptations
- Hosted as a UserControl (`CountriesNowSpace_UserControl`) that can appear in both a floating Form and a Custom Task Pane
- **Add-in export mode** is switched on by `CountriesNowSpace_UserControl.EnableAddinExportMode(hostType)` (forwarded by `CountriesNowSpace_Form.EnableAddinExportMode`). Each add-in calls it with its own host — `"Word"`/`"Excel"`/`"PowerPoint"` — from both the task-pane wrapper (`CountriesNow_TaskPane_UserControl`) and the ribbon `OnOpenCountriesForm`. In this mode:
  - **Export Type** is locked to the host Office application and **Export Library** is locked to **NetOffice**; both combos, the file-path box, the Browse button, and Clear Log are hidden — **only the Export button remains** in the export row.
  - **Export** creates a **new document in the running host application** (via `MsOfficeExportManager__Service.ExportToRunningAppAsync`) and leaves it open, unsaved — no file is written and no separate Office instance is launched. The running app is acquired on the control's own UI/STA thread via `WatcherComHelper.GetActiveCom` so the COM proxy is valid there.
  - Add-in mode is a **runtime** flag, not a compile constant — the UserControl is compiled once into `AL.WinFormsLib._netF` and referenced by the add-ins via `ProjectReference`.

### VSTO Add-in (VstoAddIn, net10.0) adaptations
- No file path, Export Type, or Export Library selection — always uses the active document
- Opened as a floating `CountriesNowSpace_WebView_Form` (BlazorWebView) — no Custom Task Pane (comhost.dll cannot activate UserControl subclasses; see repo memory `dotnet10-comhost-usercontrol-taskpane.md`)

### Blazor Web Add-in (BlazorWebAddIn) adaptations
- **Save as JSON** button — triggers a browser file download of the loaded data as a `.json` file (instead of exporting to an Office format)
- No Export Type, Export Library, or file path selection

### Blazor (AL.BlazorLib — CountriesNow__Page.razor)
- Export section added to the page: Export Type dropdown, Export Library dropdown, Export File path TextBox, Export button, Clear Log button
- Log output styled as a terminal panel (black background, green monospace text)
- File path defaults on init (and refreshes on Type/Library/DB-provider change, unless the user typed a custom path) to the **full path in the user's Documents folder** — `{MyDocuments}\ETW_CountriesNow-{library}.{ext}` (no datetime stamp), via `MsOfficeExportName_Helper.BuildFileName`. In the browser (WASM) `Environment.GetFolderPath` returns `""`, so only the file name is shown/used — `BrowserExport_AppService` downloads by file name (`Path.GetFileName`) regardless.
- The File Path TextBox uses `flex:1` so it **stretches** as the panel widens, keeping the Export / Clear Log buttons pinned to the right.
- Export writes to the specified local file path (WinForms/MAUI host has file system access); the browser host downloads the generated file.

### After Export
After a successful **file** export (non-add-in path), `CountriesNowSpace_UserControl.OpenWatcherForm()` opens the appropriate Watcher form and loads the newly created file, and `OpenExportedFile()` opens the file. The **add-in export path** (`RunAddinExportAsync`, used when `EnableAddinExportMode` was called) does neither — the new document is already open in the host application and there is no file. (A legacy `#if !MSOFFICE_ADDIN` guard remains around `OpenWatcherForm`; since the add-ins now reference `AL.WinFormsLib._netF` by `ProjectReference` rather than file-links, the UserControl compiles without `MSOFFICE_ADDIN` and the runtime `_addinExportMode` flag is what actually gates behavior.)

---

## Watcher Form Design

The Watcher form monitors Office events and optionally opens/connects to a running Office document. Its design is consistent across modalities:

### Common Controls
- **Open Files** ComboBox / list — shows currently open documents in the target Office application; selecting an item pre-fills the file path
- **File Path** TextBox + **Browse** button — specifies the document to open (standalone forms only)
- **Open / Close** button — opens or closes the target document
- **Connect / Disconnect** button — connects to or disconnects from a running instance to receive live events
- **Events** grid — lists available Office events with a checkbox to enable logging per event
- **Log** text area — displays logged event activity
- **Save JSON** button — writes the open document to a structured JSON file. The default output name follows the active document's file name with `.json` appended (e.g. `filename.docx.json`), falling back to `Documents\ETW_Ms{Host}.json` for an unsaved document. Cloud (AutoSave/OneDrive) documents with an `https` path are mapped to the local OneDrive sync folder by `OneDriveLocalPath_Helper` (Documents-folder fallback when no local copy exists). Both the net10 WebView watchers **and** the `._netF` WinForms watcher forms recompute this default **whenever the active document changes** (`SaveAsJson_Helper.BuildDefaultPath`); a path the user has typed manually is preserved.
- **Clear Log** button

### WinForms variants (AL.WinFormApp, AL.WinFormApp._netF)
- `Ms{Word|Excel|PowerPoint}_Watcher_Form` — standalone WinForms window (lives in `AL.WinFormApp._netF`)
- `Ms{Word|Excel|PowerPoint}_Watcher_WebView_Form` — BlazorWebView-hosted Blazor component (lives in `AL.WinFormApp`, net10.0 only)

#### JSON Write Method picklist (WinForms footer panel)

The footer panel of each WinForms Watcher Form includes a **JSON Write Method** ComboBox (`cbxJsonWriteMethod`), ordered `Direct` / `Interop` / `Dynamic` / `NetOffice` / `OpenXml`:

| Value | Implementation |
|-------|----------------|
| `Interop` | Calls `MsX_JsonWriter.WriteInterop(document, filePath, log)` — unwraps the live NetOffice object to its raw COM object, casts it to the `Microsoft.Office.Interop.{Host}` interface, and delegates to `DL.MsOfficeApi.Interop_Impl/JsonWriters/MsXJsonWriter` (strongly-typed Interop PIA object model) |
| `NetOffice` | Calls `MsX_JsonWriter.WriteNetOffice(document, filePath, log)` — delegates to `DL.MsOfficeApi.NetOffice_Impl/JsonWriters/MsXJsonWriter` (strongly-typed NetOffice COM object model) |
| `OpenXml` | Calls `MsX_JsonWriter.WriteOpenXml(sourceFilePath, outputJsonFilePath, log)` — delegates to `DL.MsOfficeApi.OpenXml_Impl/JsonWriters/MsXJsonWriter` (reads the file via DocumentFormat.OpenXml SDK, no running Office instance needed) |
| `Dynamic` | Calls `MsX_JsonWriter.WriteDynamic(document, filePath, log)` — unwraps the NetOffice object to its raw COM object and delegates to `DL.MsOfficeApi.Dynamic_Impl/JsonWriters/MsXJsonWriter` (late-bound `dynamic` COM, same path as the Dynamic export repos) |
| `Direct` | Calls `DL.MsOfficeApi.Direct_Impl.MsOfficeDirectVbaRunner.RunWriteMacro(app, "MSO_Ms{Host}.WriteActive{Document|Workbook|Presentation}ToJsonFile", path, log)` — runs the path-only VBA shim inside the running Office host via late-bound `Application.Run`. The macro is addressed by **`Module.Procedure`** (no `"Template.{ext}!"` qualifier): a document/template qualifier fails with `DISP_E_MEMBERNOTFOUND` when the loaded template/workbook/presentation name differs, so the macro is resolved by the loaded VBA project instead. |

All paths emit the canonical `JBC.ExploreTheWorld.DL.Ms{Host}` entity schema (see the DL section), so the watcher output matches the VBA and web add-in writers.

Each writer class is in `AL.WinFormApp/_Forms/_Watcher/`:

| Class | Source file |
|-------|-------------|
| `MsPowerPoint_JsonWriter` | `MsPowerPoint_JsonWriter.cs` |
| `MsWord_JsonWriter` | `MsWord_JsonWriter.cs` |
| `MsExcel_JsonWriter` | `MsExcel_JsonWriter.cs` |

All three classes live in `AL.WinFormsLib/_Watcher/` and are shared (via `ProjectReference`) to `AL.WinFormsLib._netF` and the VSTO projects. The `JsonWriteMethod` enum (`Direct / Interop / Dynamic / NetOffice / OpenXml`) is in `_Export/JsonWriteMethod_Enum.cs` and is likewise shared. The net10 Blazor watcher components offer the same picklist (`OpenXml` is hidden in Office add-in hosts, where the active document must stay open).

**Processing tracking:** like the net10 Blazor watchers, each `._netF` WinForms watcher wraps the write in a `Stopwatch` and logs the start and elapsed duration — `Save As JSON started ({method}) → {path}` before the write and `✔ Save As JSON complete ({method}) in {duration} → {path}` (or `✘ … failed after {duration}: …`) after. Durations format through `CL.Duration_Helper` as hours/minutes/seconds with the milliseconds as one decimal on the seconds (`45.7s`, `2m 10.0s`, `1h 2m 0.0s`).

### Console apps (net10.0-windows, JBC.ExploreTheWorld.sln only)

Two console front-ends expose the same functionality as the WinForms UI, with command line options mirroring the corresponding forms. Each has its own `MsOfficeExportRepoFactory` (the host-side factory pattern — the one place per app that instantiates the platform-specific DL repos) and delegates to the BL `MsOfficeExportManager__Service`.

| Project | Mirrors | Options |
|---------|---------|---------|
| `AL.ExportData.ConsoleApp` | Countries API form export | `--type Word\|Excel\|PowerPoint` · `--method Interop\|Dynamic\|NetOffice\|OpenXml` · `--provider SqliteDb\|SqlServerDb\|InMemoryDb\|AccessDb` · `--file <path>` (defaults `ETW_CountriesNow.{ext}`). Loads countries through the DB-backed manager (check DB → API fallback → persist) with flag-image enrichment; connection strings come from an optional `appsettings.json`. |
| `AL.SaveAsJson.ConsoleApp` | Watcher forms Save-As-JSON | `--source <document>` (required) · `--method Direct\|Interop\|Dynamic\|NetOffice\|OpenXml` · `--type Word\|Excel\|PowerPoint` (default: inferred from the source extension) · `--output <path>` (default: source name + `.json`). `Direct` opens the document in its Office host and runs `MSO_Ms{Host}.WriteActive…ToJsonFile` — it requires the ExploreTheWorld VBA macros to be loaded in the host (routed through the `DL.MsOfficeApi.Direct_Impl` repos, like every other method). |

### VSTO Add-in variants
- `._netF` addins host the UserControl in both a floating Form and a Custom Task Pane
- `net10.0` addins open floating Blazor (WebView) forms only — no task panes (`comhost.dll` cannot activate `UserControl` subclasses; see repo memory `dotnet10-comhost-usercontrol-taskpane.md`)

### Blazor Web Add-in
- `Watcher.razor` page in the WASM client monitors Office.js events for the active document
- No Connect/Disconnect selection (Office.js API abstracts the host)

### Blazor (AL.BlazorLib — Watcher Pages)
- `MsWord_Watcher__Page.razor`, `MsExcel_Watcher__Page.razor`, `MsPowerPoint_Watcher__Page.razor` — full Blazor UI (not stubs)
- Display is driven by the injected `WatcherEvent_AppService` singleton, which the WinForms host pushes events into
- Open Files dropdown, file path input, Save as JSON, Clear Log, terminal-style log panel (black/green)
- Output path defaults to the active document's name with the selected write method inserted before the extension and `.json` appended (e.g. `Report-NetOffice.pptx.json`), via `SaveAsJson_Helper.BuildDefaultPath(current, "ETW_Ms{Host}", method)`, recomputed when the active document or the selected write method changes
- `WatcherEvent_AppService` registered as `services.AddSingleton<WatcherEvent_AppService>()` in `AL.WinFormApp/Program.cs`

### Access VBA Watcher Forms
- `MsWord`, `MsExcel`, `MsPowerPoint` forms in the Access database
- Include a ComboBox of currently open files (populated by Timer every 3 seconds)
- Follow the same Open / Close / Connect pattern using WithEvents class modules

#### Access Export Method picklist

Each Watcher form includes a `cmb_ExportMethod` ComboBox with two options:

| Value | Implementation |
|-------|----------------|
| `COM` | Calls `MSO_MsX_JsonWriter.WriteXxxToJsonFile(oDocument, sPath)` directly from Access (early-bound Office reference) |
| `Direct` | Calls `oApp.Run "ExploreTheWorld.{ext}!MSO_MsX.WriteActiveXToJsonFile", sPath` — delegates to the shim inside the running macro add-in |

The old `ETW__MsX_JsonWriter` modules have been removed. All three hosts (PPT/Word/Excel) now use the `MSO_MsX_JsonWriter` pattern consistently.

---

### AL.MsOffice{Word|Excel|PowerPoint}BlazorWebAddIn + .Client — Office Web Add-ins

**Purpose:** Demonstrate Office.js-based task pane add-ins using the current hybrid Blazor host + client layout on net10.0.

**Architecture:**
- Server project (net10.0) + Blazor WebAssembly client project (net10.0) paired per host
- Local Office debugging assets live in the server project (`package.json`, `Assets/manifest.local.xml`)
- **SharedRuntime** required: `ExecuteFunction` (Ribbon button) and task pane share one JS runtime
- **Toggle pattern:** `commands.js` uses `Office.addin.showAsTaskpane()` / `Office.addin.hide()` — NOT `ShowTaskpane` action (cannot hide)
- **JS interop:** per-page ES module imported in `OnAfterRenderAsync(firstRender)` via `IJSObjectReference`
- **Events:** JS registers Office.js event handlers, calls back via `DotNet.invokeMethodAsync` to `[JSInvokable] OnEventLogged`
- **Icon images:** `wwwroot/Images/Button{16,32,64,80}x{16,32,64,80}.png` referenced from manifest

**Event scope per host:**
| Host | Events available |
|------|-----------------|
| Word | Selection via common API `Office.EventType.DocumentSelectionChanged` (no production `Word.Document.onSelectionChanged` exists); `document.onParagraphAdded/Changed/Deleted` + `onContentControlAdded` (document-level); `ContentControl.onDeleted/onEntered/onExited` registered per existing control instance (WordApi 1.5) |
| Excel | `worksheets.onActivated/onDeactivated/onAdded/onDeleted` (ExcelApi 1.7); `worksheets.onSelectionChanged` (ExcelApi 1.9, any sheet) with fallback to the active sheet's `onSelectionChanged` (1.7) |
| PowerPoint | Selection via common API `Office.EventType.DocumentSelectionChanged`; `presentation.onSlideSelectionChanged` is preview-only (PowerPointApi BETA) and is skipped with a console warning on production hosts |

**Hybrid project structure (per host):**
```
AL.MsOffice{Host}BlazorWebAddIn/
  ExploreTheWorld.AL.MsOffice{Host}BlazorWebAddIn.csproj  # Server host (net10.0)
  package.json
  Assets/
    manifest.local.xml
  Components/
    _Imports.razor                 # Includes @using for .Client.Components (enables RenderModeInfo in layout)
    Layout/
      MainLayout.razor             # RadzenLayout with RenderModeInfo WASM island in header
      MainLayout.razor.cs          # Trademark/framework properties
    Pages/                         # Task pane pages and shared UI
  AL.MsOffice{Host}BlazorWebAddIn.Client/
    ExploreTheWorld.AL.MsOffice{Host}BlazorWebAddIn.Client.csproj  # WASM client (net10.0)
    Program.cs
    Components/
      RenderModeInfo.razor         # renderModeBadge + breakpointBadge; WASM interactive island
    Pages/
      Home.razor + .cs + .js
      DocumentInfo.razor + .cs
      Events.razor + .cs
```

**Ribbon tab "ETW (Web)" (`Assets/manifest.local.xml`) — 3 groups, one button each:**

| Group | Button | Task pane route |
|-------|--------|-----------------|
| Countries API | Countries Pane (Blazor) | `/countries-now` |
| Watcher | Watcher Pane (Blazor) | `/events` |
| Export | Save as JSON (Blazor) | `/save-as-json` |

Each button is a `ShowTaskpane` control pointing the shared runtime at the route above.

**`RenderModeInfo.razor` pattern** — shared across all three `.Client` projects. Renders as a `@rendermode InteractiveWebAssembly` island when included in the static-SSR server layout header. Shows:
- `renderModeBadge` (`BadgeStyle.Success` for Server, `BadgeStyle.Info` for WebAssembly)
- `breakpointBadge` (XS→Danger, SM→Warning, MD→Info, LG→Success, XL→Primary, XXL→Secondary)

The breakpoint is resolved once via `window.getWindowWidth()` in `OnAfterRenderAsync(firstRender)`. The function is defined in each add-in's `App.razor` as a global script before `blazor.web.js`.

**SDK / packages:** server host uses `Microsoft.NET.Sdk.Web`; client uses `Microsoft.NET.Sdk.BlazorWebAssembly`.

**Manifest project:** `net481`, SDK `Microsoft.NET.Sdk` — holds only the XML manifest and `SharePointProjectItem.spdata` (sideload helper). Build output is copied to `_sideload/{Host}/Web/`.

**Ports:**
| Host | IIS Express | Kestrel HTTPS | Kestrel HTTP |
|------|-------------|---------------|--------------|
| Word | 44332 | 7097 | 5097 |
| Excel | 44333 | 7098 | 5098 |
| PowerPoint | 44334 | 7099 | 5099 |

**Manifest IDs:**
| Host | Id |
|------|----|
| Word | `A1B2C3D4-E5F6-7890-ABCD-EF1234567890` |
| Excel | `B2C3D4E5-F6A7-8901-BCDE-F12345678901` |
| PowerPoint | `C3D4E5F6-A7B8-9012-CDEF-012345678902` |

---

### AL.MsOffice{Word|Excel|PowerPoint}VstoAddIn — COM-Hosted VSTO-Style Add-ins (net10.0, WebView only)

**Purpose:** Demonstrate COM-hosted .NET 10 Office add-ins using NetOffice__10 project references. Each add-in opens floating Blazor (WebView) forms from `AL.WinFormApp` via shared-link compilation.

**SDK Type:** `Microsoft.NET.Sdk.WindowsDesktop`  
**Framework:** `net10.0-windows`, `UseWindowsForms=true`, `EnableComHosting=true`

**Key Characteristics:**
- `Addin` class implements `IDTExtensibility2` + `IRibbonExtensibility` (no `ICustomTaskPaneConsumer`; no task panes)
- Ribbon XML loaded from embedded resource via `Assembly.GetManifestResourceStream`
- `OnConnection` wraps the native `Application` object: `new Word.Application(null, Application)`
- Ribbon tab **"ETW (VSTO)"** — 3 groups:

  | Group | Button | imageMso | Callback |
  |-------|--------|----------|----------|
  | Countries API | Countries Form (Blazor) | `ViewsFormView` | `OnOpenCountriesNowBlazorForm` → opens `CountriesNowSpace_WebView_Form` |
  | Watcher | Watcher Form (Blazor) | `ReviewReviewingPaneVertical` | `OnOpenWatcherBlazorForm` → opens `Ms{Host}_Watcher_WebView_Form` |
  | Export | Save as JSON | `FileSaveAs` | `OnSaveAsJson` (C# NetOffice writer) |
  | Export | Save as JSON (Direct) | `FileSaveAs` | `OnSaveAsJsonVba` → runs the `MSO_Ms{Host}.WriteActive…` VBA macro |
- `RemoveLegacyTaskPaneProgIds()` called in `OnConnection` to clean up any stale `HKCU\Software\Classes\{ProgId}` entries from older versions that used task panes
- `[ComRegisterFunction]` / `[ComUnregisterFunction]` write `HKCU\Software\Microsoft\Office\{Host}\Addins\{ProgId}` registry keys (LoadBehavior=3, FriendlyName, Description) and the KB 948461 lockback bypass key
- CLSID registered to `comhost.dll` (not `mscoree.dll`) in `InprocServer32`

**Addin CLSIDs (net10.0):**
| Host | Addin CLSID |
|------|-------------|
| Word | `B1A2C3D4-E5F6-4789-ABCD-EF0123456789` |
| Excel | `C2B3D4E5-F6A7-4890-BCDE-F01234567890` |
| PowerPoint | `D3C4E5F6-A7B8-4901-CDEF-012345678901` |

**Shared-link files (from `AL.WinFormApp/_Forms/`):**
- `_Export/`: `ExportMenuHelper.cs`, `ExportMethod_Enum.cs`, `ExportType_Enum.cs`, `JsonWriteMethod_Enum.cs`
- `_Watcher/`: `MsOfficeEvent_Record.cs`, `MsOfficeEvents_Repo.cs`, `MsOfficeJsonWriter_Helper.cs`, `WatcherComHelper.cs`, `Ms{Host}_JsonWriter.cs`
- WebView forms: `CountriesNowSpace_WebView_Form.cs/.Designer.cs`, `Ms{Host}_Watcher_WebView_Form.cs/.Designer.cs`

**NetOffice__10 project references:** `NetOffice.Analyzers`, `NetOffice`, `OfficeApi`, plus host-specific `WordApi` / `ExcelApi` / `PowerPointApi`.

**Build target `RegisterComAddin64`** (Debug only): writes all required `HKCU` registry keys automatically after build.

---

### AL.MsOffice{Word|Excel|PowerPoint}VstoAddIn._netF — VSTO-Style Add-ins (net481)

**Purpose:** .NET Framework 4.8.1 equivalent using `NetOfficeFw.*` NuGet packages and NetOffice's `COMAddin` base class.

**SDK Type:** `Microsoft.NET.Sdk`  
**Framework:** `net481`, `UseWindowsForms=true`

**Key Characteristics:**
- `Addin` class inherits `COMAddin` (from `NetOffice.{Host}Api.Tools`)
- Class decorated with `[COMAddin]`, `[ProgId]`, `[Guid]`, `[CustomUI]` attributes — **no `[CustomPane]`** (see per-window panes below)
- **Per-window Custom Task Panes (Word/Excel/PowerPoint are all SDI).** A CTP belongs to exactly one document window, and NetOffice's `[CustomPane]` creates a single pane bound to whichever window was active at add-in load — so it only ever appears on that first document. Instead the add-in keeps **no `[CustomPane]`** and creates a pane per window on demand: `CTPFactoryAvailable` calls `base` (to build `TaskPaneFactory`) but creates nothing; the toggle handlers call `GetOrCreatePane(...)`, which looks up two `Dictionary<int hwnd, CustomTaskPane>` maps (watcher / countries) keyed by the active window's `Hwnd` (PowerPoint: `DocumentWindow.HWND`), and on a miss calls `TaskPaneFactory.CreateCTP(progId, title, window.UnderlyingObject)`, sets `DockPosition=Right`/`Width`, wires `VisibleStateChangeEvent`, and (watcher only) casts `pane.ContentControl` to inject the running `Application`. A dead entry (window closed) is detected by a throwing `Visible` read and recreated.
- `CustomUI_OnLoad` override calls `base` (which sets the `RibbonUI` property used to refresh the toggle buttons)
- `OnError(ErrorMethodKind, Exception)` override + `NetOffice.DebugConsole.Default` routed to `%TEMP%\…._netF.netoffice.log` — captures the COM exception from a failed `CreateCTP`, which the base `COMAddin` otherwise only writes to its internal `DebugConsole`
- `OnOpenWatcher` creates a floating `Form` (940×660) containing `Ms{Host}_Watcher_UserControl`
- `OnOpenCountriesForm` opens `CountriesNowSpace_Form` on a new STA thread
- Toggle callbacks: `OnCheckPanelToggle`/`OnGetPressedPanelToggle` (Watcher pane) and `OnCheckCountriesPane`/`OnGetPressedCountriesPane` (Countries pane); each resolves the **active window's** pane and reads/sets its `Visible`
- Each created pane's `VisibleStateChangeEvent` → `OnPaneVisibleStateChanged` invalidates both toggle buttons (fires for the pane's own `[X]` too)
- `OnSaveAsJson` / `OnSaveAsJsonVba`: JSON export via C# writer or VBA macro shim

**Ribbon tab "ETW (VSTO._netF)" — 3 groups:**

| Group | Button | imageMso | Callback |
|-------|--------|----------|----------|
| Countries API | Countries Form | `ViewsFormView` | `OnOpenCountriesForm` |
| Countries API | Countries Pane toggle | `ReviewReviewingPaneVertical` | `OnCheckCountriesPane` |
| Watcher | Watcher Form | `ViewsFormView` | `OnOpenWatcher` |
| Watcher | Watcher Pane toggle | `ReviewReviewingPaneVertical` | `OnCheckPanelToggle` |
| Export | Save as JSON | `FileSaveAs` | `OnSaveAsJson` |
| Export | Save as JSON (Direct) | `FileSaveAs` | `OnSaveAsJsonVba` |

**UserControl files:**
- `Ms{Host}_Watcher_UserControl.cs` — the CTP-hosted watcher control. **Lives in `AL.WinFormsLib._netF`** (referenced by `ProjectReference`), not in the add-in project — so its COM registration must point at that assembly (see the watcher-pane registration note below). Implements `ITaskPane` (its `OnConnection` injects the running `Application` into the embedded watcher form, so the pane is live once docked) and `IObjectSafety` (see the ActiveX-safety note below).
- `CountriesNow_TaskPane_UserControl.cs` + `.Designer.cs` — **in the add-in project**; wrapper with parameterless ctor; creates `CountriesNowSpaceManager__Service` manually (no DI); hosts `CountriesNowSpace_UserControl` dock-filled inside and calls `EnableAddinExportMode(host)`; implements `ITaskPane` and `IObjectSafety`. Needed because `ICTPFactory.CreateCTP` requires a parameterless constructor and `CountriesNowSpace_UserControl` has none in the `NETFRAMEWORK` build.

**Shared-link files from `AL.WinFormApp/_Forms/_Watcher/` (non-form helpers):**
- `MsOfficeEvent_Record.cs`, `MsOfficeEvents_Repo.cs`, `MsOfficeJsonWriter_Helper.cs`, `WatcherComHelper.cs`, `Ms{Word|Excel|PowerPoint}_JsonWriter.cs`, `JsonWriteMethod_Enum.cs`

**Shared-link files from `AL.WinFormApp._netF/_Forms/` (Countries form + UC):**
- `CountriesNowSpace_Form.cs`, `CountriesNowSpace_Form.Designer.cs`, `CountriesNowSpace_Form.resx`
- `CountriesNowSpace_UserControl.cs`, `CountriesNowSpace_UserControl.Designer.cs`

**`MSOFFICE_ADDIN` preprocessor symbol:** Defined in the Word `._netF.csproj`. It only affects files compiled **inside** the add-in project. Since `CountriesNowSpace_UserControl` is now compiled once in `AL.WinFormsLib._netF` (referenced via `ProjectReference`, not file-linked into the add-in), its `#if !MSOFFICE_ADDIN` guard around `OpenWatcherForm` is always compiled in; the add-in vs. desktop distinction is made at **runtime** via `EnableAddinExportMode` / the `_addinExportMode` flag instead.

**Packages:** `NetOfficeFw.Core` 1.9.10, `NetOfficeFw.{Host}` 1.9.10, `Microsoft.CSharp` 4.7.0, `System.Net.Http.Json` 10.0.9

**Addin CLSIDs (net481):**
| Host | Addin CLSID | Watcher UC CLSID | Countries Pane UC CLSID |
|------|-------------|------------------|-------------------------|
| Word | `E4D5E6F7-B8C9-4A12-DEF0-123456789012` | `A1B2C3D4-E5F6-7890-ABCD-EF0123456702` | `C1D2E3F4-A5B6-7890-ABCD-EF0123456710` |
| Excel | `F5E6F7A8-C9DA-4B23-EF01-234567890123` | `A1B2C3D4-E5F6-7890-ABCD-EF0123456704` | `C2D3E4F5-A6B7-7890-ABCD-EF0123456711` |
| PowerPoint | `A6F7A8B9-D0E1-4C34-F012-345678901234` | `A1B2C3D4-E5F6-7890-ABCD-EF0123456706` | `C3D4E5F6-A7B8-7890-ABCD-EF0123456712` |

**Build target `RegisterComAddin`** (Debug only): registers the Addin CLSID (via `mscoree.dll`) and **both** UserControl CLSIDs (Watcher + Countries API pane) needed for `ICTPFactory.CreateCTP`. Each UserControl CLSID also gets the `Implemented Categories` keys `{7DD95801-…}` (SafeForScripting) and `{7DD95802-…}` (SafeForInitializing) — see the ActiveX-safety note below.

> **ActiveX-safety shim is required for net481 (not net10).** Office hosts a Custom Task Pane's content through its **ActiveX container**, which queries the control for `IObjectSafety` before hosting it. A hand-registered .NET `UserControl` that answers neither `IObjectSafety` nor the `SafeForScripting`/`SafeForInitializing` component categories is treated as untrusted, and `ICTPFactory.CreateCTP` **fails silently** — the `TaskPaneInfo` stays in `TaskPanes` with `IsLoaded=false`/`Pane=null`, so the ribbon toggle sets `TaskPaneInfo.Visible` into a dictionary that is never applied and **no pane ever appears**. Every net481 CTP-hosted control therefore implements `IObjectSafety` (declared once in `AL.WinFormsLib._netF/IObjectSafety.cs`, IID `CB5BDC81-…`, returning safe-for-caller+data), and `RegisterComAddin` writes the two component-category keys as belt-and-suspenders. The `#if !NETFRAMEWORK`-guarded `ICustomQueryInterface` shim (which hides `IProvideClassInfo`/`IProvideClassInfo2`) is a **net10-only** workaround for a different failure (HRESULT `0x80131165`) and is not needed on net481.

> **Watcher-pane registration must target the shared library.** The `Ms{Host}_Watcher_UserControl` types live in `AL.WinFormsLib._netF` (namespace `JBC.ExploreTheWorld.AL.WinFormsLib`), not in the add-in project. NetOffice asks Office to `CreateCTP` using the type's default ProgId (`Namespace.ClassName`), so the registry entries **must** use that ProgId, the matching `[Guid]` CLSID (`A1B2C3D4-…-0270{2|4|6}`), `Class = JBC.ExploreTheWorld.AL.WinFormsLib.Ms{Host}_Watcher_UserControl`, `Assembly = JBC.ExploreTheWorld.AL.WinFormsLib._netF`, and `CodeBase = $(TargetDir)JBC.ExploreTheWorld.AL.WinFormsLib._netF.dll`. Registering the old add-in-namespace ProgId/class (a leftover from when the control lived in the add-in project) makes `CreateCTP` fail silently: the **Watcher pane never registers**, `TaskPanes` shifts by one, and the fixed-index toggle handlers (`TaskPanes[0]` = watcher, `[1]` = countries) then drive the wrong pane — the classic "Watcher pane won't open, Countries pane opens from the wrong button" symptom. The Countries pane is unaffected because its `CreateCTP` type (`CountriesNow_TaskPane_UserControl`) genuinely lives in the add-in assembly.

**net10.0 COM hosting limitation:** `comhost.dll` cannot activate a `UserControl` subclass via `CoCreateInstance` (returns `CLASS_E_CLASSNOTAVAILABLE`). The `._netF` `COMAddin` base class handles task pane creation automatically via `NetOfficeFw`. See repo memory `dotnet10-comhost-usercontrol-taskpane.md` for the full investigation and workaround.

---

## Country Slides — Reveal.js Integration

The `/country-slides` route (`CountrySlides__Page`) renders all loaded countries as a reveal.js slide presentation, providing a rich interactive view of the CountriesNow data.

### Static Assets

Reveal.js 5.x files (core + all official plugins + all themes) are kept in:

```
DL.MsJSInterop.RevealJs/wwwroot/revealjs/
  reset.css, reveal.css, reveal.esm.js, reveal.js
  theme/  (black, white, league, beige, sky, night, serif, simple, solarized, blood, moon, dracula, ...)
  plugin/highlight/  (highlight.esm.js, monokai.css, zenburn.css)
  plugin/notes/      (notes.esm.js — speaker notes window)
  plugin/search/     (search.esm.js — in-slide full-text search)
  plugin/zoom/       (zoom.esm.js — alt-click zoom)
  plugin/math/       (math.esm.js, katex.js, mathjax2.js, mathjax3.js)
  plugin/markdown/   (markdown.esm.js)
```

Because they live in the RCL's `wwwroot/`, they are served at:
```
/_content/JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs/revealjs/…
```
in both the server-hosted (`AL.BlazorWebApp`) and standalone WASM (`AL.BlazorWebApp.ClientOnly`) deployments — no duplication.

All reveal CSS (`reset.css`, `reveal.css`, `plugin/highlight/monokai.css`, and the active theme) is injected into `document.head` by `reveal-interop.js` during `initialize` and removed by `destroy`. No `<HeadContent>`/`HeadOutlet` is involved, so the slides work in every host — including BlazorWebView (WinForms/MAUI) and Oqtane, which have no `HeadOutlet`. The `etw-slide-*` layout styles live in `AL.BlazorLib/wwwroot/css/country-slides.css` (`@import`ed by `app.css`; referenced directly as an Oqtane module resource).

### JavaScript Interop Layer

**`DL.MsJSInterop.RevealJs/wwwroot/js/reveal-interop.js`** — scoped ES module; lazy-loaded by the interop service.

Key exports:

| Export | Description |
|---|---|
| `initialize(el, dotNetRef, transition, theme, logLevel)` | Injects the reveal CSS links, creates `Reveal` instance with Zoom, Notes, Search, Highlight plugins; wires `slidechanged`, `overviewshown/hidden` events back to Blazor. Returns `true` on success. |
| `navigateNext/Prev/Right/Left/Up/Down()` | Slide navigation |
| `navigateToSlide(h, v)` | Jump to absolute position |
| `setTransition(name)` | Live transition change via `reveal.configure()` |
| `setTheme(name)` | Swaps `#revealjs-theme` `<link>` href in `<head>` to the named bundled theme |
| `startAutoPlay(intervalMs)` / `stopAutoPlay()` | Timer-driven auto-advance with wrap |
| `toggleOverview()` / `togglePause()` | Pass-through to reveal API |
| `requestFullscreen(el)` / `isFullscreen()` | Fullscreen API wrapper |
| `layout()` | Force reveal layout recalculation |
| `destroy()` | Cleans up reveal instance, injected CSS links, ResizeObserver, fullscreen listener |

**Path resolution** — `reveal-interop.js` resolves the reveal.js library and CSS relative to its own URL (`new URL('../revealjs/', import.meta.url)`), so assets load correctly regardless of the host's base href.

**`DL.MsJSInterop.RevealJs/RevealJs/RevealJs__Interop.cs`** — implements `RevealJs__Interop__Interface` (formerly `CountrySlides__Repo` / `CountrySlides__Repo__Interface`); derives `JsModuleInterop__Base` for the lazy-load-with-SemaphoreSlim-lock + cache-bust pattern. Module path:

```csharp
"./_content/JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs/js/reveal-interop.js"
```

### Blazor Page — `CountrySlides__Page`

**Route:** `/country-slides`  
**Namespace:** `JBC.ExploreTheWorld.AL.BlazorLib.Countries`  
**Base class:** `CountrySlides__PageBase : ComponentBase, IAsyncDisposable`

**Slide structure (rendered as Razor markup):**

| Slide | Content |
|---|---|
| Title | Gradient background, globe emoji, country count, keyboard hint |
| Per-country (×N) | Unique HSL background derived from ISO2, flag image (cached PNG via `FlagImageManager__Service`, emoji fallback), country name, ISO2/ISO3 badges — each element a sequential `fragment fade-up` |
| End | Green gradient, "World Tour Complete!" count |

**Country slide background color** — deterministic from ISO2:
```csharp
var hash = iso2.Aggregate(0, (h, c) => h * 31 + c);
return $"hsl({Math.Abs(hash) % 360}, 50%, 20%)";
```

**Flag images** — after reveal.js initializes, the component background-loads real flag PNGs
through the optionally-injected `FlagImageManager__Service` (local cache first, Wikimedia
download fallback) and renders them as `data:image/png;base64` `<img>` elements
(`.etw-slide-flag-img`), batching `StateHasChanged` every 20 images. Hosts without a registered
flag image manager — and countries whose image cannot be resolved — keep the emoji flag.

**Flag emoji fallback** — computed from ISO 3166-1 alpha-2 Regional Indicator codepoints:
```csharp
string.Concat(iso2.ToUpperInvariant()
    .Select(c => char.ConvertFromUtf32(0x1F1E6 + (c - 'A'))));
```

**Toolbar controls (Radzen):**

| Control | Component | Function |
|---|---|---|
| Theme | `RadzenDropDown` | Swaps reveal CSS theme live |
| Transition | `RadzenDropDown` | `none / slide / convex / concave / zoom / fade` |
| ← / Slide N/Total / → | `RadzenButton` + `RadzenText` | Navigation |
| Auto-play speed | `RadzenDropDown` | 2 / 4 / 6 / 10 sec |
| Play/Pause | `RadzenButton` | Timer-driven auto-advance |
| Overview | `RadzenButton` | Toggle reveal overview mode (O key) |
| Slide numbers | `RadzenButton` | Toggle `c/t` slide number overlay |
| Fullscreen | `RadzenButton` | Fullscreen API |
| Progress | `RadzenProgressBar` | Mirrors reveal slide position (H axis) |

**Jump-bar** — a horizontal scrollable strip of `RadzenButton` chips (flag + ISO2) beneath the presentation; clicking jumps directly to that country's slide. The active slide button is highlighted with `ButtonStyle.Primary`.

**JS callbacks (JSInvokable):**

| Method | Trigger |
|---|---|
| `OnSlideChanged(h, v, total)` | Every `slidechanged` event |
| `OnOverviewChanged(bool)` | `overviewshown` / `overviewhidden` |
| `OnFullscreenChanged(bool)` | `fullscreenchange` DOM event |

### DI Registration

Both `AL.BlazorWebApp/Program.cs` and `AL.BlazorWebApp.ClientOnly/Program.cs`:

```csharp
builder.Services.AddScoped<RevealJs__Interop__Interface, RevealJs__Interop>();
```

### Navigation

`AL.BlazorLib/_Shared/NavMenu.razor` includes:
```razor
<RadzenPanelMenuItem Text="Country Slides" Icon="slideshow" Path="/country-slides" />
```

### Plugins Used

| Plugin | Purpose |
|---|---|
| `RevealZoom` | Alt-click to zoom into slide content |
| `RevealNotes` | Speaker notes (press S to open separate window) |
| `RevealSearch` | Ctrl+Shift+F full-text search across all slides |
| `RevealHighlight` | Syntax highlighting (monokai theme) for code blocks |

### Extension Points for Third-Party Plugins

The `reveal-interop.js` module is structured so additional ESM plugins can be added to the `plugins: [...]` array in `initialize()`. Candidates from the [reveal.js wiki](https://github.com/hakimel/reveal.js/wiki/Plugins,-Tools-and-Hardware):

- **Chart.js integration** (`reveal-chart`) — population bar charts per-slide
- **Countdown** — auto-advance with visual countdown timer
- **Anything** — embed arbitrary HTML/SVG/Canvas per slide
- **D3.js** — animated data visualisations per country

---

## Oqtane Module Architecture

ExploreTheWorld integrates with the **Oqtane framework** (v10.2.1, net10.0) as a modular CMS host. Three modules and a Radzen-based site theme are provided under `src/AL.Oqtane/`.

### Assembly Naming Requirement

Oqtane's `AssemblyExtensions.IsOqtaneAssembly()` only discovers assemblies whose filename contains `"oqtane"` (case-insensitive). All ETW module assembly names include `.AL.Oqtane.` to satisfy this requirement:

| Module | Client AssemblyName | Server AssemblyName |
|--------|---------------------|---------------------|
| Radzen | `JBC.ExploreTheWorld.AL.Oqtane.Radzen__Module.Client` | `JBC.ExploreTheWorld.AL.Oqtane.Radzen__Module.Server` |
| CountriesNow | `JBC.ExploreTheWorld.AL.Oqtane.CountriesNow__Module.Client` | `JBC.ExploreTheWorld.AL.Oqtane.CountriesNow__Module.Server` |
| CountrySlides | `JBC.ExploreTheWorld.AL.Oqtane.CountrySlides__Module.Client` | `JBC.ExploreTheWorld.AL.Oqtane.CountrySlides__Module.Server` |

### Module Pair Pattern (Client + Server)

Each module is split into two projects:

- **Client** (`Microsoft.NET.Sdk.Razor`): contains `IModule` (ModuleInfo.cs), `ModuleBase`-derived page components (Index.razor), and `_Imports.razor`. Adds `Private="false"` ProjectReferences to `Oqtane.Shared` and `Oqtane.Client` to prevent redundant DLL copying.
- **Server** (`Microsoft.NET.Sdk`): contains `IClientStartup` (Startup/ClientStartup.cs) for DI registration. ProjectReferences its Client counterpart (without Private=false) plus all ETW service stack projects.

**DLL propagation chain:** `Oqtane.Server → {Module}.Server → {Module}.Client → Oqtane.Client` (Private=false). The transitive chain ensures Client DLLs appear in `Oqtane.Server/bin`, where Oqtane's assembly scanner picks them up.

### `IClientStartup` Instead of `IServerStartup`

`IServerStartup` is defined in `Oqtane.Server`, not `Oqtane.Shared`. Referencing it from a module Server project would create a circular dependency (`Oqtane.Server` → module → `Oqtane.Server`). ETW modules implement `IClientStartup` (from `Oqtane.Shared`) instead. In Blazor Server mode, Oqtane calls `IClientStartup.ConfigureServices` on the server — so all scoped services are correctly registered.

### RenderMode Override for Interactive Components

Oqtane's default render mode is `"Static"` (set in `appsettings.json`). Components using JSInterop (Radzen services, Reveal.js) require interactive rendering. Each module's `Index.razor` overrides this:

```razor
@inherits ModuleBase

<CountriesNow__Component />

@code {
    public override string RenderMode => RenderModes.Interactive;
}
```

### Module Descriptions

#### `Radzen__Module` — Site-wide Radzen Theme + Components

Registers Radzen's four scoped services (`DialogService`, `NotificationService`, `TooltipService`, `ContextMenuService`) and injects `_content/Radzen.Blazor/css/default.css` as a site-wide stylesheet via `ModuleDefinition.Resources`. Both Client and Server reference `Radzen.Blazor` NuGet 11.0.5.

```
src/AL.Oqtane/
  Radzen__Module.Client/
    ModuleInfo.cs          # IModule, Resources: [Stylesheet("_content/Radzen.Blazor/css/default.css")]
    _Imports.razor
  Radzen__Module.Server/
    Startup/ClientStartup.cs   # IClientStartup: registers 4 Radzen scoped services
```

#### `CountriesNow__Module` — Countries Grid + Browser Export

Hosts `CountriesNow__Component` from `AL.BlazorLib`. Registers the full ETW service stack (EF Core `IDbContextFactory`, `CountriesNowSpaceApiManager__Repo`, HTTP client for the API, BL manager, `FileDownload__Interop`, `BrowserExport_AppService`, and a `DbProvider_AppService` singleton with `ProviderName = "SqlServerDb"`). `download-file.js` is no longer a global `ModuleDefinition.Resources` script — `BrowserExport_AppService` injects `FileDownload__Interop__Interface`, which lazy-imports the ESM module from `DL.MsJSInterop`.

> **Every `[Inject]` dependency of a hosted component must be registered in `ClientStartup`.** A missing registration (e.g. `DbProvider_AppService`, which `CountriesNow__Component` injects) does not just break that module — the property-injection failure crashes the whole Blazor Server circuit, and its scoped services are disposed while other components are mid-request. The visible symptom is often a misleading downstream error such as `ObjectDisposedException: Cannot access a disposed object ('System.Net.Http.HttpClient')` thrown from Oqtane's own `ControlPanelInteractive`/`ServiceBase`, and every other interactive module on the site (e.g. Country Slides) rendering empty until the page is reloaded.

Like the slideshow module, `Index.razor` passes an explicit viewport-based height (the same `calc(100vh - 240px)` as `CountrySlides__Module`). `CountriesNow__Component` is a full-height flex column whose grid area scrolls internally (`flex:1; overflow:auto`) while the export log and options stay pinned to the bottom (`flex-shrink:0`). Without a bounded height the flex layout collapses in Oqtane's auto-height pane: the grids stop scrolling internally and the whole page scrolls instead, pushing the export log and options off the bottom.

```razor
<CountriesNow__Component Style="height:calc(100vh - 240px);min-height:480px;" />
```

```
src/AL.Oqtane/
  CountriesNow__Module.Client/
    ModuleInfo.cs          # Radzen resources only (no download-file.js script)
    Index.razor            # RenderModes.Interactive + explicit height Style
    _Imports.razor
  CountriesNow__Module.Server/
    Startup/ClientStartup.cs   # IClientStartup: DbContextFactory + Repos + FileDownload__Interop + BrowserExport_AppService
```

#### `CountrySlides__Module` — Reveal.js Country Slideshow

Hosts `CountrySlides__Component` from `AL.BlazorLib`. Registers the same ETW service stack as CountriesNow plus `RevealJs__Interop` (scoped JS interop service for Reveal.js, from `DL.MsJSInterop.RevealJs`).

Oqtane panes are auto-height (unlike the web app shell, where `CountrySlides__Page` passes `Style="height:100%"` inside a full-height `RadzenBody`), so `Index.razor` passes an explicit viewport-based height — otherwise the reveal.js viewport collapses to 0 px and only the toolbar/jump bar render:

```razor
<CountrySlides__Component Style="height:calc(100vh - 240px);min-height:480px;" />
```

```
src/AL.Oqtane/
  CountrySlides__Module.Client/
    ModuleInfo.cs
    Index.razor            # RenderModes.Interactive + explicit height Style
    _Imports.razor
  CountrySlides__Module.Server/
    Startup/ClientStartup.cs   # IClientStartup: DbContextFactory + Repos + RevealJs__Interop
```

### ExploreTheWorld Theme (`Theme/`)

`ExploreTheWorld.AL.Oqtane.Theme` (assembly `JBC.ExploreTheWorld.AL.Oqtane.Theme`) is a Radzen-based Oqtane theme that mirrors the Blazor web app shell: navigation in a left sidebar, Oqtane's search and login controls in the header.

```
src/AL.Oqtane/
  Theme/
    ThemeInfo.cs               # ITheme: bootstrap CSS+JS + Radzen.Blazor.js + theme.js resources
    ThemeSettings.razor        # ISettingsControl: Radzen theme CSS picker, Login/Register toggles
    Layouts/Default.razor      # RadzenLayout: header + sidebar + body panes
    Containers/Container.razor # Default module container
    Controls/SidebarMenu.razor # MenuBase-derived RadzenPanelMenu (left navigation)
    wwwroot/css/theme.css      # Header/sidebar color sync (Radzen CSS variable overrides)
    wwwroot/js/theme.js        # Sidebar toggle (plain JS — the theme renders statically)
    _Imports.razor
```

- **The theme renders statically.** The site render mode is `Static`; only the ETW modules override to Interactive. Two consequences: Blazor `@onclick` handlers in the layout never fire (the sidebar toggle is plain JS in `wwwroot/js/theme.js`, loaded as a theme Script resource), and `<HeadContent>` never renders because **Oqtane has no `HeadOutlet`** — head content must go through Oqtane's resource pipeline or `ThemeBase.AddHeadContent`.
- **`Layouts/Default.razor`** — `RadzenLayout` with a `RadzenHeader` (sidebar toggle, `Logo`, spacer, then Oqtane's `Search`, `UserProfile`, `Login`, and `ControlPanel` controls), a `RadzenSidebar` hosting `SidebarMenu`, and a `RadzenBody` with the Default, Admin, and Footer panes. The sidebar uses `Responsive="false"` — responsive mode never applies the `rz-sidebar-expanded`/`rz-sidebar-collapsed` classes that `theme.js` toggles. `<RadzenComponents />` is rendered at the end of the layout as the host for Radzen's scoped UI services.
- **Radzen stylesheet loading** — the baseline `default.css` is declared in the layout's `Resources` override (evaluated via `Activator.CreateInstance` during page composition, so the list must be static — no `PageState` access). When the `RadzenCss` site setting differs from the default, `OnParametersSet` emits an overriding link via `AddHeadContent`, which renders after the page resource links and wins the cascade. Legacy `*-base.css` setting values are normalized to the complete theme files in code and by `SQL/004`.
- **`Controls/SidebarMenu.razor`** — inherits Oqtane's `MenuBase` (same page filtering/permission logic as the built-in menus) and renders a nested `RadzenPanelMenu`. Well-known ETW page paths map to the same Material icons the web app `NavMenu` uses (`home`, `public`, `slideshow`); Oqtane CSS-class icons (`"oi oi-..."`) are skipped because Radzen icons are Material Symbols names.
- **`ThemeInfo.cs` resources** — bootstrap CSS + JS (still required by the Oqtane built-in controls the theme hosts — Search, Login, ControlPanel offcanvas — and by the Oqtane admin modules), the theme's own `theme.css` (header background/text redefined to the sidebar variables, so the two stay in sync in every selectable Radzen theme) and `theme.js`, plus `_content/Radzen.Blazor/Radzen.Blazor.js`. Loading `Radzen.Blazor.js` at theme level is load-bearing: module-level `Script` resources are not injected into the initial document for interactive modules, so `RadzenDataGrid` fails with `Could not find 'Radzen.createDataGrid' ('Radzen' was undefined)` when the page relies on the module resource alone.
- **Radzen CSS** — all references use the complete theme files (`default.css`, `dark.css`, `material.css`, ...) rather than the `-base` variants, per the repo UI-styling rules.
- **Site default** — `SQL/004_Set_ETW_Theme.sql` sets `Site.DefaultThemeType` / `DefaultContainerType` to this theme's `Default` layout and `Container`. The theme registers itself in the `[Theme]` table on startup; the script only switches the site defaults.

### Database Connection

All ETW modules share the same SQL Server LocalDB instance as Oqtane (`DefaultConnection` in `appsettings.json`). The EF Core factory is registered with a runtime `IConfiguration` lookup to avoid compile-time coupling:

```csharp
services.AddDbContextFactory<ExploreTheWorldDbContext>((sp, options) =>
{
    var config = sp.GetRequiredService<IConfiguration>();
    options.UseSqlServer(config.GetConnectionString("DefaultConnection")!);
}, ServiceLifetime.Transient);
```

EF Core creates ETW tables automatically via `EnsureCreated()` on first use. Numbered idempotent SQL scripts (run against the Oqtane tenant database) live in `src/AL.Oqtane/SQL/`:

| Script | Purpose |
|--------|---------|
| `001_Create_ETW_Tables.sql` | Manual ETW schema creation / inspection reference |
| `002_Create_ETW_Pages.sql` | Inserts the Countries Now and Country Slides pages + module instances |
| `003_Fix_ETW_Page_Permissions.sql` | Re-points permission rows at the correct Oqtane role names |
| `004_Set_ETW_Theme.sql` | Sets the ExploreTheWorld Radzen theme as the site default theme/container |

### Oqtane.Server.csproj Integration

The three Server modules are referenced from `oqtane.framework/Oqtane.Server/Oqtane.Server.csproj`. This pulls the full module DLL tree into `Oqtane.Server/bin` at build time — no module hot-loading or copy step required during development.

---

## Testing Architecture

Eleven net10.0 test projects are included in `JBC.ExploreTheWorld.sln` under the **Tests** solution folder (plus the `._netF` counterparts in `JBC.ExploreTheWorld._netF.sln`). All net10.0 test projects use **xUnit 2.9.3**, **FluentAssertions 8.x**, and **coverlet.collector**.

### Test Framework Stack Summary

| Project | Framework | Data Access | Mocking | Browser/UI |
|---|---|---|---|---|
| **UnitTests** | xUnit, Moq, FluentAssertions | None (mocked) | Moq | — |
| **IntegrationTests** | xUnit, FluentAssertions | In-Memory EF Core | — | — |
| **OpenXmlLibTests** | xUnit, FluentAssertions | Temp `.docx`/`.xlsx`/`.pptx` files | — | — |
| **RazorTests** | bUnit, Moq, FluentAssertions | None (mocked) | Moq | bUnit (JSDOM) |
| **WebAppTests** | xUnit, Playwright, FluentAssertions | — | — | Playwright (Chromium) |
| **OqtaneTests** | xUnit, Playwright, FluentAssertions | Oqtane LocalDB (read-only) | — | Playwright (Chromium) |
| **WinFormAppTests** | xUnit, FlaUI, FluentAssertions | — | — | FlaUI UIA3 |
| **MauiAppTests** | xUnit, FlaUI, FluentAssertions | — | — | FlaUI UIA3 |
| **OfficeAddinTests** | xUnit, FlaUI, NetOffice__10, FluentAssertions | — | — | FlaUI UIA3 |
| **OfficeWebAddinTests** | xUnit, Playwright, FluentAssertions | — | — | Playwright (Chromium) |
| **AccessDbTests** | xUnit, FlaUI, FluentAssertions | `VBA\Access\ExploreTheWorld.accdb` (read-only) | — | FlaUI UIA3 (MSACCESS.EXE) |

`SqliteDbTests._netF` (net481 only) additionally covers the `DL.CountriesNowSpaceData._netF` repository against a real SQLite database file via the EF Core 3.x provider.

**UI test screenshots**: every UI test project writes `before.png`/`after.png` under `TestResults/{Project}/{Class}/{Test}/`. The test bases bring the target window to the foreground and wait for a settle delay (`ETW_TEST_SETTLE_MS`, defaults 1500–4000 ms per project) before capturing — WebView2/Blazor content paints asynchronously, so capturing immediately after launch produces blank images.

---

### ExploreTheWorld.UnitTests

**Path:** `src/Tests/UnitTests/ExploreTheWorld.UnitTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0`

**Purpose:** Tests where no actual data access happens. All repositories and APIs are mocked with Moq.

**Key Packages:** `Moq 4.20.72`

**Project References:** CL, DL, DL.CountriesNowSpaceApi, BL, AL

**Test Coverage:**
- `CL/_Row_Tests.cs` — `_Row` entity base class behaviour
- `CL/DataResult_Row_Tests.cs` — `DataResult_Row<T>` and `DataSource_Enum`
- `BL/CountriesNowSpaceManager_Tests.cs` — BL manager DB-first / API-fallback logic with mocked dependencies

**Pattern:**
```csharp
public class CountriesNowSpaceManager_Tests
{
    private readonly Mock<CountriesNowSpaceApi_Interface> _mockApi;
    private readonly Mock<CountriesNowSpaceApiManager__Repo__Interface> _mockDbManager;
    private readonly CountriesNowSpaceManager__Service _sut;
    // ...
}
```

---

### ExploreTheWorld.IntegrationTests

**Path:** `src/Tests/IntegrationTests/ExploreTheWorld.IntegrationTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0`

**Purpose:** Tests where actual data access happens against an in-memory database. No mocking of repositories.

**Key Packages:** `Microsoft.EntityFrameworkCore.InMemory 10.0.9`, `Microsoft.AspNetCore.Mvc.Testing 10.0.9`

**Project References:** CL, DL, DL.CountriesNowSpaceApi, DL.CountriesNowSpaceData, BL, AL

**Fixture:** `Fixtures/DatabaseFixture.cs` creates an isolated in-memory database per test class; `Fixtures/TestDbContextFactory.cs` wraps options for the `IDbContextFactory<T>` interface the repo requires.

**Test Coverage:**
- `DL/CountriesNowSpaceApiManager_Tests.cs` — `CountriesNowSpaceApiManager__Repo` CRUD operations

**Pattern:**
```csharp
public class CountriesNowSpaceApiManager_Tests : IClassFixture<DatabaseFixture>
{
    public CountriesNowSpaceApiManager_Tests(DatabaseFixture fixture)
    {
        _sut = new CountriesNowSpaceApiManager__Repo(fixture.CreateFactory());
    }
}
```

---

### ExploreTheWorld.OpenXmlLibTests

**Path:** `src/Tests/OpenXmlLibTests/ExploreTheWorld.OpenXmlLibTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0` (`OpenXmlLibTests._netF` links all source for `net481`)

**Purpose:** Tests the `DL.MsOfficeApi.OpenXml_Impl` document-export repositories against real files: each test exports to a temp `.docx`/`.xlsx`/`.pptx`, re-opens it with the OpenXml SDK, and verifies the structure, then deletes the file.

**Project References:** CL, DL, DL.MsOfficeApi.OpenXml_Impl

**Test Coverage:**
- `Managers/MsWord_OpenXml__Repo_Tests.cs` — Word export table structure, log progress, `WriteDocumentJsonAsync`
- `Managers/MsExcel_OpenXml__Repo_Tests.cs` — Countries worksheet rows, log progress, `WriteDocumentJsonAsync`
- `Managers/MsPowerPoint_OpenXml__Repo_Tests.cs` — title + data slides, log progress, `WriteDocumentJsonAsync`

---

### ExploreTheWorld.SqliteDbTests._netF

**Path:** `src/Tests/SqliteDbTests._netF/ExploreTheWorld.SqliteDbTests._netF.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net481` (standalone — no net10.0 counterpart; the net10.0 `IntegrationTests` cover the repository against EF Core InMemory)

**Purpose:** Tests the `DL.CountriesNowSpaceData._netF` repository against a real SQLite database file using the EF Core 3.x Sqlite provider (`CreateDbContext()` only).

**Fixture:** `Fixtures/SqliteDatabaseFixture.cs` creates a seeded temp `.db` file per test class and deletes it on disposal. `ClearAllAsync` runs in its own test class so its own fixture database is cleared, never the shared one.

---

### ExploreTheWorld.RazorTests

**Path:** `src/Tests/RazorTests/ExploreTheWorld.RazorTests.csproj`  
**SDK:** `Microsoft.NET.Sdk.Razor` | **Target:** `net10.0`

**Purpose:** Blazor component tests using bUnit. Radzen services are pre-registered in `BlazorTestBase`; all data services are mocked.

**Key Packages:** `bunit 2.0.33-preview`, `Moq 4.20.72`

**Project References:** AL.BlazorLib, AL, BL, CL, DL, DL.CountriesNowSpaceApi, DL.MsJSInterop, DL.MsSystemNet

**Base class:** `BlazorTestBase : BunitContext` — registers Radzen services (`DialogService`, `NotificationService`, `TooltipService`, `ContextMenuService`) and `WatcherEvent_AppService`; sets `JSInterop.Mode = Loose`.

**Test Coverage:**
- `Countries/CountriesNow_Page_Tests.cs` — `CountriesNow__Page` renders without error and contains expected controls

---

### ExploreTheWorld.WebAppTests

**Path:** `src/Tests/WebAppTests/ExploreTheWorld.WebAppTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0`

**Purpose:** End-to-end Playwright tests against a running `AL.BlazorWebApp` instance. Scripts are in C#.

**Key Packages:** `Microsoft.Playwright 1.52.0`

**Prerequisites:** `AL.BlazorWebApp` must be running. Set `ETW_BASE_URL` env var (default: `https://localhost:7000`).

**Base class:** `PlaywrightTestBase : IAsyncLifetime` — launches headless Chromium, creates a page, disposes on teardown.

**Test Coverage:**
- `Countries/CountriesNow_Tests.cs` — loads Countries page, verifies Load button, triggers data fetch
- `Navigation/Navigation_Tests.cs` — verifies home and country-slides routes respond

**To install Playwright browsers:** `pwsh bin/Debug/net10.0/playwright.ps1 install`

---

### ExploreTheWorld.OqtaneTests

**Path:** `src/Tests/OqtaneTests/ExploreTheWorld.OqtaneTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0`

**Purpose:** End-to-end Playwright tests for the custom Oqtane modules and the ExploreTheWorld Radzen theme, against the local Oqtane host.

**Key Packages:** `Microsoft.Playwright 1.61.0`

**Prerequisites:** `oqtane.framework/Oqtane.Server` must be built, and the local site installed (LocalDB `Oqtane-ETW`, ETW pages from `SQL/002`, ETW theme from `SQL/004`). The fixture starts `Oqtane.Server.exe` if it is not already listening. Env vars: `ETW_OQTANE_URL` (default `http://localhost:44357`), `ETW_OQTANE_PATH` (default auto-discovered).

**Base class / fixture:** `PlaywrightTestBase : IAsyncLifetime` (also captures browser console errors into `ConsoleErrors` — Blazor circuit failures surface there, not as HTTP errors) + `OqtaneServerFixture` collection fixture.

**Test Coverage:**
- `Theme/Theme_Tests.cs` — RadzenLayout shell, left sidebar `RadzenPanelMenu`, search/login in header, Radzen stylesheet loaded via theme resources (no `-base.css`), JS sidebar toggle collapse/expand
- `Countries/CountriesNow_Tests.cs` — module renders without DI property-injection errors, grid auto-loads rows, `Radzen.Blazor.js` initialized before the grid
- `Countries/CountrySlides_Tests.cs` — reveal.js slides render, viewport has non-zero height (Oqtane auto-height pane regression), toolbar country-count badge
- `Navigation/Navigation_Tests.cs` — countries → slides client-side navigation keeps the shared Blazor Server circuit alive (regression guard for the `DbProvider_AppService` DI crash)

**To install Playwright browsers:** `pwsh bin/Debug/net10.0/playwright.ps1 install`

---

### ExploreTheWorld.WinFormAppTests

**Path:** `src/Tests/WinFormAppTests/ExploreTheWorld.WinFormAppTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0-windows`

**Purpose:** Windows UI Automation tests against a running `AL.WinFormApp` instance using FlaUI.

**Key Packages:** `FlaUI.Core 4.0.0`, `FlaUI.UIA3 4.0.0`

**Prerequisites:** `AL.WinFormApp` must be built. Set `ETW_WINFORMAPP_PATH` env var or the base class auto-discovers the Debug/Release bin path.

**Base class:** `FlaUITestBase : IDisposable` — launches the app, exposes `Application`, `UIA3Automation`, and `MainWindow`; disposes on teardown.

**Test Coverage:**
- `Forms/MainForm_Tests.cs` — app launches, shows main window with correct title and Countries Now button

---

### ExploreTheWorld.MauiAppTests

**Path:** `src/Tests/MauiAppTests/ExploreTheWorld.MauiAppTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0-windows`

**Purpose:** Windows UI Automation tests against a running `AL.MauiApp.WinUI` instance (unpackaged MAUI Blazor exe) using FlaUI.

**Key Packages:** `FlaUI.Core 5.0.0`, `FlaUI.UIA3 5.0.0`

**Prerequisites:** `AL.MauiApp.WinUI` must be built. Set `ETW_MAUIAPP_PATH` env var or the base class auto-discovers `AL.MauiApp.WinUI\bin\[x64\]{Debug|Release}\net10.0-windows10.0.19041.0\win-x64\JBC.ExploreTheWorld.AL.MauiApp.WinUI.exe`.

**Base class:** `FlaUITestBase : IDisposable` — launches the app, exposes `Application`, `UIA3Automation`, and `MainWindow`; disposes on teardown.

**Test Coverage:**
- `Pages/MainPage_Tests.cs` — app launches, main window has the expected title and hosts a BlazorWebView (WebView2)

---

### ExploreTheWorld.OfficeAddinTests

**Path:** `src/Tests/OfficeAddinTests/ExploreTheWorld.OfficeAddinTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0-windows`

**Purpose:** Windows UI Automation tests against a running Office application with the ETW VSTO add-in loaded. Uses FlaUI for UI Automation and NetOffice__10 (external project references) for COM interop.

**Key Packages:** `FlaUI.Core 4.0.0`, `FlaUI.UIA3 4.0.0`, `NetOfficeFw.Core 1.9.10`, `NetOfficeFw.Word/Excel/PowerPoint 1.9.10`

**External Project References:** `NetOffice`, `OfficeApi`, `WordApi`, `ExcelApi`, `PowerPointApi` from `code-zgh-NetOfficeFw__NetOffice__10`

**Prerequisites:** Word/Excel/PowerPoint must be running with the ETW VSTO add-in loaded. Tests skip gracefully if the Office host is not running.

**Base class:** `FlaUIOfficeTestBase : IDisposable` — attaches to a running Office process by name; tests skip automatically if the process is not found.

**Test Coverage:**
- `Word/MsWordVstoAddin_Tests.cs` — ETW VSTO ribbon tab visible, Countries Form (Blazor) button present
- `Excel/MsExcelVstoAddin_Tests.cs` — ETW VSTO ribbon tab visible
- `PowerPoint/MsPowerPointVstoAddin_Tests.cs` — ETW VSTO ribbon tab visible

---

### ExploreTheWorld.OfficeWebAddinTests

**Path:** `src/Tests/OfficeWebAddinTests/ExploreTheWorld.OfficeWebAddinTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0`

**Purpose:** End-to-end Playwright tests against a running Office Web Add-in server (not through Office Desktop). Scripts are in C#.

**Key Packages:** `Microsoft.Playwright 1.52.0`

**Prerequisites:** The add-in servers must be running. Set per-host env vars:
- `ETW_WORD_ADDIN_URL` (default: `https://localhost:7097`)
- `ETW_EXCEL_ADDIN_URL` (default: `https://localhost:7098`)
- `ETW_PPT_ADDIN_URL` (default: `https://localhost:7099`)

**Base class:** `PlaywrightOfficeTestBase : IAsyncLifetime` — launches headless Chromium with HTTPS errors ignored; exposes per-host base URLs.

**Test Coverage:**
- `Word/MsWordWebAddin_Tests.cs` — home, document-info, events pages load successfully
- `Excel/MsExcelWebAddin_Tests.cs` — home and events pages load successfully
- `PowerPoint/MsPowerPointWebAddin_Tests.cs` — home and events pages load successfully

---

### ExploreTheWorld.AccessDbTests

**Path:** `src/Tests/AccessDbTests/ExploreTheWorld.AccessDbTests.csproj`  
**SDK:** `Microsoft.NET.Sdk` | **Target:** `net10.0-windows`

**Purpose:** UI-automation tests that launch Microsoft Access with `VBA\Access\ExploreTheWorld.accdb` and perform user actions with FlaUI (open forms from the Navigation Pane, verify controls).

**Key Packages:** `FlaUI.Core 5.0.0`, `FlaUI.UIA3 5.0.0`

**Prerequisites:** Microsoft Access installed (tests return early when `MSACCESS.EXE` or the `.accdb` is missing). Env overrides: `ETW_MSACCESS_PATH`, `ETW_ACCESS_DB_PATH`.

**Base class:** `FlaUIAccessTestBase : IDisposable` — launches `MSACCESS.EXE "<db>" /ro` (read-only, so the compact-on-close setting never rewrites the tracked `.accdb`), waits for the UI to settle, captures screenshots, and closes/kills Access on disposal.

**Test Coverage:**
- `Database/AccessApp_Tests.cs` — database opens, Navigation Pane lists the `cns_Country` and `Ms*` forms
- `Forms/CountriesNowDataForm_Tests.cs` — the `cns_Country` form ("CountriesNow Data") opens from the Navigation Pane and shows the Load / Export / Browse... controls

---

### Running Tests

```powershell
# All test projects in the prescribed order (both frameworks)
.\Scripts\Run-AllTests.ps1

# Unit tests only
dotnet test src\Tests\UnitTests\ExploreTheWorld.UnitTests.csproj

# Integration tests only
dotnet test src\Tests\IntegrationTests\ExploreTheWorld.IntegrationTests.csproj

# OpenXml document export tests
dotnet test src\Tests\OpenXmlLibTests\ExploreTheWorld.OpenXmlLibTests.csproj

# SQLite (_netF, EF Core 3.x) data tests
dotnet test src\Tests\SqliteDbTests._netF\ExploreTheWorld.SqliteDbTests._netF.csproj

# Razor (bUnit) tests only
dotnet test src\Tests\RazorTests\ExploreTheWorld.RazorTests.csproj

# Web app tests (requires running AL.BlazorWebApp)
dotnet test src\Tests\WebAppTests\ExploreTheWorld.WebAppTests.csproj

# Oqtane module/theme tests (requires built Oqtane.Server + installed local site)
dotnet test src\Tests\OqtaneTests\ExploreTheWorld.OqtaneTests.csproj

# WinForms app tests (requires built AL.WinFormApp, Windows only)
dotnet test src\Tests\WinFormAppTests\ExploreTheWorld.WinFormAppTests.csproj

# Office add-in tests (requires Word/Excel/PowerPoint running with ETW add-in, Windows only)
dotnet test src\Tests\OfficeAddinTests\ExploreTheWorld.OfficeAddinTests.csproj

# Office web add-in tests (requires add-in servers running)
dotnet test src\Tests\OfficeWebAddinTests\ExploreTheWorld.OfficeWebAddinTests.csproj

# Access database UI tests (requires Microsoft Access, Windows only)
dotnet test src\Tests\AccessDbTests\ExploreTheWorld.AccessDbTests.csproj
```

