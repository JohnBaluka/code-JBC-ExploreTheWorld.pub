# File Structure & Organization Guide

## Overview

The ExploreTheWorld codebase follows a consistent folder structure based on architectural layers and domain concerns. This guide provides the folder conventions and file organization patterns used across all projects.

## Solution-Level Organization

```
code-JBC-ExploreTheWorld/
├── src/                                  # All .NET projects, the four .sln files, and Tests/
│   ├── JBC.ExploreTheWorld.sln              # Main solution (all projects)
│   ├── JBC.ExploreTheWorld.AL.BlazorLib._radzen.sln
│   ├── JBC.ExploreTheWorld.AL.BlazorWebApp.sln
│   ├── JBC.ExploreTheWorld._netF.sln
│   │
│   ├── CL/                                   # Common Layer (net10.0)
│   ├── CL._netF/                             # Common Layer (.NET Framework)
│   ├── DL/                                   # Dependency Layer (net10.0)
│   ├── DL._netF/                             # Dependency Layer (.NET Framework)
│   ├── DL.MsJSInterop/                       # Generic JS interop (FileDownload/Layout) + JsModuleInterop__Base
│   ├── DL.MsJSInterop.RevealJs/              # Reveal.js interop (RevealJs__Interop) + revealjs library
│   ├── DL.MsOfficeApi.MsOfficeJs.Word_Impl/        # Word Office.js page interops
│   ├── DL.MsOfficeApi.MsOfficeJs.Excel_Impl/       # Excel Office.js page interops
│   ├── DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl/  # PowerPoint Office.js page interops
│   ├── DL.MsSystem/                          # System Utilities
│   ├── DL.MsSystemNet/                       # Network Utilities
│   ├── DL.CountriesNowSpaceApi/              # CountriesNow API client
│   ├── DL.CountriesNowSpaceApi._netF/        # CountriesNow API (.NET Framework)
│   ├── DL.MsOfficeApi.OpenXml_Impl/                        # OpenXML export repos (Word/Excel/PowerPoint → file)
│   ├── DL.MsOfficeApi.OpenXml_Impl._netF/                  # OpenXML export repos (.NET Framework, shared-link)
│   ├── DL.MsOfficeApi.NetOffice_Impl/                      # NetOffice COM export repos (net10.0-windows)
│   ├── DL.MsOfficeApi.NetOffice_Impl._netF/                # NetOffice export repos (.NET Framework, NetOfficeFw NuGet)
│   ├── DL.MsOfficeApi.Dynamic_Impl/                        # COM late-binding (dynamic) export repos (net10.0-windows)
│   ├── DL.MsOfficeApi.Dynamic_Impl._netF/                  # COM late-binding repos (.NET Framework, shared-link)
│   ├── DL.MsOfficeApi.Direct_Impl/                         # VBA Application.Run JSON-writer repos (net10.0-windows)
│   ├── DL.MsOfficeApi.Direct_Impl._netF/                   # VBA Application.Run repos (.NET Framework, shared-link)
│   ├── BL/                                   # Business Logic Layer (net10.0)
│   ├── BL._netF/                             # Business Logic Layer (.NET Framework)
│   ├── AL/                                   # Application Layer (net10.0)
│   ├── AL._netF/                             # Application Layer (.NET Framework)
│   ├── AL.BlazorLib/                         # Reusable Blazor Components (Razor SDK)
│   ├── AL.BlazorLib._radzen/                 # Radzen-Focused Components
│   ├── AL.BlazorLib.Server._radzen/          # Server-Side Blazor + Radzen
│   ├── AL.BlazorWebApp/                      # Hybrid Web App Server
│   ├── AL.BlazorWebApp.Client/               # WebAssembly Client (companion to AL.BlazorWebApp)
│   ├── AL.BlazorWebApp.ClientOnly/           # Standalone WASM PWA (no server required)
│   ├── AL.WinFormApp/                        # WinForms BlazorWebView host (net10.0)
│   │   └── _Forms/IObjectSafety.cs          # ActiveX safety shim used by task-pane controls
│   ├── AL.WinFormApp._netF/                  # WinForms traditional app (.NET Framework)
│   │
│   ├── AL.ExportData.ConsoleApp/             # Console export (same options as the Countries API form)
│   ├── AL.SaveAsJson.ConsoleApp/             # Console Save-As-JSON (same options as the Watcher forms)
│   │
│   ├── AL.MsOfficeWordVstoAddIn/             # Word COM add-in (net10.0, comhost)
│   ├── AL.MsOfficeWordVstoAddIn._netF/       # Word COM add-in (.NET Framework, COMAddin base)
│   ├── AL.MsOfficeExcelVstoAddIn/            # Excel COM add-in (net10.0, comhost)
│   ├── AL.MsOfficeExcelVstoAddIn._netF/      # Excel COM add-in (.NET Framework, COMAddin base)
│   ├── AL.MsOfficePowerPointVstoAddIn/       # PowerPoint COM add-in (net10.0, comhost)
│   ├── AL.MsOfficePowerPointVstoAddIn._netF/ # PowerPoint COM add-in (.NET Framework, COMAddin base)
│   │
│   ├── AL.MsOfficeWordBlazorWebAddIn/        # Word Office.js host add-in (server + manifest)
│   ├── AL.MsOfficeWordBlazorWebAddIn.Client/ # Word Blazor client components
│   ├── AL.MsOfficeExcelBlazorWebAddIn/       # Excel Office.js host add-in (server + manifest)
│   ├── AL.MsOfficeExcelBlazorWebAddIn.Client/# Excel Blazor client components
│   ├── AL.MsOfficePowerPointBlazorWebAddIn/  # PowerPoint Office.js host add-in (server + manifest)
│   ├── AL.MsOfficePowerPointBlazorWebAddIn.Client/ # PowerPoint Blazor client components
│   ├── AL.Oqtane/                            # Oqtane modules + Radzen theme (was Oqtane.ExploreTheWorld/)
│   └── Tests/                                # All 17 test projects (see docs/architecture.md → Testing Architecture)
│
├── docs/                                 # Markdown docs (this file, architecture.md, ...) + screenshots/
├── TestResults/                          # UI-test before.png/after.png output ({Project}/{Class}/{Test}/)
├── slides/                               # Slide JSON exports (HowToThinkInBlazor-*.pptx.json)
├── Scripts/, SQL/, VBA/, TestData/, Samples/, oqtane.framework/   # unchanged
└── README.md
```

---

## Layer Folder Structure

### Common Layer (CL)

**Purpose:** Shared utilities and models

**Folder Pattern:**
```
CL/
├── ExploreTheWorld.CL.csproj
├── Enum/
│   └── Enum_Extensions.cs               # Enum helper methods
├── TreeStructure/
│   └── TreeNode.cs                      # Hierarchical tree structure
├── _Row/
│   ├── _Row.cs                          # Base entity model
│   ├── RowLog_Row.cs                    # Row audit log
│   └── ColumnLog_Row.cs                 # Column audit log
├── _Services/
│   └── ILogger2.cs                      # Custom logger interface
└── bin/, obj/                           # Build artifacts (auto-generated)
```

**File Naming:**
- `_Row.cs` suffix for domain entities
- `_Extensions.cs` for static extension methods
- `__Repo__Interface` suffix for repo interfaces; `__Service__Interface` suffix for service interfaces

---

### Dependency Layer (DL)

> "DL" = **Dependency Layer** (external dependencies wrapped behind interfaces: DB, HTTP, system, JS interop), not "Data Layer".

**Purpose:** Data persistence and query abstraction

**Folder Pattern:**
```
DL/
├── ExploreTheWorld.DL.csproj
├── Managers/
│   ├── [Feature]Manager.cs              # Data access managers
│   └── FlagImageManager/                # Flag image cache abstractions
│       ├── FlagImageStore__Repo__Interface.cs     # cache store contract (file system / IndexedDB)
│       ├── FlagImageStore_FileSystem__Repo.cs     # %LocalAppData%\JBC\ExploreTheWorld\FlagImages cache
│       ├── FlagImageDownload__Repo__Interface.cs  # image download contract
│       └── _Rows/FlagImage_Row.cs                 # resolved image (bytes + path + source)
├── Models/
│   └── [Feature]_Model.cs               # Entity models (inherit _Row)
├── Repositories/
│   ├── [Feature]__Repo__Interface.cs    # Data access contracts
│   └── [Feature]__Repo.cs               # Repository implementations
├── MsOfficeApi/                         # Office API contracts + canonical Save-As-JSON schema
│   ├── Ms{Word|Excel|PowerPoint}__Repo__Interface.cs  # export/JSON-write repo contracts
│   ├── MsOfficeRunningAppExport__Repo__Interface.cs
│   ├── MsOffice/                        # Canonical Save-As-JSON infrastructure
│   │   ├── MsOfficeJsonSerializer.cs    # Single serializer all writers use
│   │   ├── MsOfficeJsonWriterOptions.cs # Blob output options (base64 | separate files)
│   │   ├── MsOfficeUndefined.cs         # "**Undefined" / -99 markers
│   │   ├── _Entities/ImageBlob.cs       # Image blob (Extension, Base64, FileName)
│   │   ├── _Fields/
│   │   └── _Enums/{TriState_Enum,BlobOutput_Enum}.cs
│   ├── MsPowerPoint/                    # Canonical PowerPoint entities (VBA object model)
│   │   ├── _Entities/                   # Presentation, Slide, Shape, ... (derive _Fields)
│   │   ├── _Fields/                     # Scalar VBA properties per entity
│   │   └── _Enums/                      # ShapeType_Enum, AutoShapeType_Enum, ... (all have Undefined = -99)
│   ├── MsExcel/                         # Canonical Excel entities (Workbook, Sheet, Cell, ...)
│   │   ├── _Entities/ + _Fields/
│   ├── MsWord/                          # Canonical Word entities (Document, Paragraph, Table, ...)
│   │   ├── _Entities/ + _Fields/
│   └── MsOfficeJs/                      # Office.js-shaped rows ({Host}…Js_Row)
│       └── {PowerPoint|Excel|Word}/     #   + Ms{Host}JsMapper (to/from the canonical entities)
├── Domain/                              # Domain-specific subfolders
│   ├── Users/
│   ├── Content/
│   └── Settings/
├── App.cs                               # Layer definition
└── bin/, obj/                           # Build artifacts
```

The `MsOfficeApi/Ms*` folders define the canonical "Save as JSON" schema shared by the VBA writers,
`DL.MsOfficeApi.NetOffice_Impl/JsonWriters/` (strongly-typed NetOffice), `DL.MsOfficeApi.Dynamic_Impl/JsonWriters/`
(late-bound `dynamic` COM), `DL.MsOfficeApi.OpenXml_Impl/JsonWriters/`, and the web add-in
`Ms{Host}JsMapper` classes.

**File Naming:**
- `Manager.cs` suffix for data managers
- `__Repo__Interface` suffix for repo interface files
- `__Repo` suffix for repository implementations
- `_Model.cs` suffix for entity models

**Domain-Specific Organization:**
Create subfolders for logical domains:
```
DL/
├── Domain/
│   ├── Users/
│   │   ├── User_Model.cs
│   │   ├── UserRepository__Repo__Interface.cs
│   │   └── UserRepository__Repo.cs
│   ├── Content/
│   │   ├── Content_Model.cs
│   │   ├── ContentRepository__Repo__Interface.cs
│   │   └── ContentRepository__Repo.cs
│   └── Viewers/
│       └── ...
```

---

### Specialized DL Projects

#### DL.MsJSInterop and DL.MsJSInterop.* - JavaScript Interop

JS interop lives in dedicated dependency projects. `.razor` components inject a typed `{Name}__Interop__Interface` (registered in `Program.cs`) instead of importing JS modules or calling `IJSRuntime` directly.

```
DL.MsJSInterop/                           # generic browser interop
├── ExploreTheWorld.DL.MsJSInterop.csproj
├── _Shared/JsModuleInterop__Base.cs      # lazy ESM-import + cache-bust base class
├── FileDownload/                         # FileDownload__Interop(__Interface)
├── FlagImageCache/                       # FlagImageCache__Interop(__Interface) — IndexedDB flag cache
│                                         #   + FlagImageStore_Browser__Repo (implements the DL store interface)
├── Layout/                               # Layout__Interop(__Interface)
└── wwwroot/js/{download-file,flag-image-cache,layout}.js  # ESM modules

DL.MsJSInterop.RevealJs/                  # reveal.js wrapper
├── RevealJs/RevealJs__Interop(__Interface).cs
└── wwwroot/{js/reveal-interop.js, revealjs/**}

DL.MsOfficeApi.MsOfficeJs.{Word|Excel|PowerPoint}_Impl/   # per-host Office.js page interops
├── {DocumentInfo|WorkbookInfo|PresentationInfo}/    # {Host}…Info__Interop + _Row
├── Events/                                          # {Host}Events__Interop
├── SaveAsJson/                                      # {Host}SaveAsJson__Interop + Result_Row + Data_Row
└── wwwroot/js/{…-info,events,save-as-json,home}.js
```

These per-host Office.js interop projects are grouped under `DL.MsOfficeApi.MsOfficeJs.*` (with the
`_Impl` suffix per the Office-repo convention) but still reference `DL.MsJSInterop` for the shared
`JsModuleInterop__Base`. The Office.js-shaped Save-As-JSON rows and `Ms{Host}JsMapper` classes live in
the core `DL/MsOfficeApi/MsOfficeJs/{Host}/` folders; the `{Host}SaveAsJson__Interop` classes consume them.

**Pattern** — derive `JsModuleInterop__Base`, which lazily imports the ESM module from `_content/{AssemblyName}/js/` and invokes its exports:
```csharp
public class FileDownload__Interop : JsModuleInterop__Base, FileDownload__Interop__Interface
{
    public FileDownload__Interop(IJSRuntime jsRuntime)
        : base(jsRuntime, "./_content/JBC.ExploreTheWorld.DL.MsJSInterop/js/download-file.js") { }

    public async Task DownloadFileFromBytesAsync(string fileName, string contentType, byte[] bytes)
    {
        var module = await GetModuleAsync();
        await module.InvokeVoidAsync("downloadFileFromBytes", fileName, contentType, bytes);
    }
}
```

#### DL.MsSystem - System Utilities

```
DL.MsSystem/
├── ExploreTheWorld.DL.MsSystem.csproj
├── System/
│   ├── FileOperations.cs                # File I/O patterns
│   └── RegistryOperations.cs            # Registry access
└── bin/, obj/
```

#### DL.MsSystemNet - Network & HTTP

```
DL.MsSystemNet/
├── ExploreTheWorld.DL.MsSystemNet.csproj
├── Http/
│   ├── HttpClientFactory.cs             # HTTP client setup
│   └── RestClient.cs                    # REST patterns
└── bin/, obj/
```

---

### Business Logic Layer (BL)

**Purpose:** Business rules and validation

**Folder Pattern:**
```
BL/
├── ExploreTheWorld.BL.csproj
├── ServiceCollectionExtensions.cs       # AddCountriesNowSpaceDbSwitcher(...)
├── _Services/                           # BL services (namespace JBC.ExploreTheWorld.BL — "_Services" is dropped)
│   ├── CountriesNowSpaceManager__Service.cs
│   ├── FlagImageManager__Service.cs
│   ├── MsOfficeExportManager__Service.cs
│   └── DbProviderSwitcher__Service.cs             # (the export-repo factory seam is a DL contract in DL/MsOfficeApi/)
├── Domain/                              # Domain-specific subfolders (when organizing by domain)
│   ├── Users/
│   │   ├── UserManager.cs
│   │   ├── UserValidator.cs
│   │   └── UserService.cs
│   ├── Content/
│   │   └── ...
│   └── Viewers/
│       └── ...
└── bin/, obj/
```

**File Naming:**
- `{Name}__Service` suffix for BL orchestrator services (concrete — no interface of their own; mock via the injected DL interfaces)
- `__Service__Interface` reserved for AL service interfaces; the host-implemented export-repo factory seam is a **DL** contract (`DL.MsOfficeApi.MsOfficeExportRepoFactory__Interface`), not a BL type
- Leading-underscore folders (`_Services/`) are grouping-only and are **not** part of the namespace

**Manager Pattern:**
```csharp
namespace JBC.ExploreTheWorld.BL.Users
{
    public class UserManager
    {
        private readonly ILogger<UserManager> _logger;
        private readonly IUserRepository _repository;

        public UserManager(ILogger<UserManager> logger, IUserRepository repository)
        {
            _logger = logger;
            _repository = repository;
        }

        public async Task<UserResult> CreateUserAsync(CreateUserRequest request)
        {
            // Validate business rules
            var validator = new UserValidator();
            var validation = validator.Validate(request);
            
            if (!validation.IsValid)
                return UserResult.Failure(validation.Errors);

            // Execute business logic
            var user = new User_Model { ... };
            await _repository.CreateAsync(user);
            
            return UserResult.Success(user);
        }
    }
}
```

---

### Application Layer (AL)

**Purpose:** Application orchestration

**Folder Pattern:**
```
AL/
├── ExploreTheWorld.AL.csproj
├── App.cs                               # Layer definition (minimal)
└── bin/, obj/
```

The core AL is minimal. Most application logic lives in specialized AL projects (AL.BlazorLib, AL.BlazorWebApp).

---

## Blazor Component Library (AL.BlazorLib)

**Purpose:** Reusable Blazor components and infrastructure

**Folder Pattern:**
```
AL.BlazorLib/
├── ExploreTheWorld.AL.BlazorLib.csproj
├── _Imports.razor                       # Global usings
├── _Shared/
│   ├── Base__RadzenComponent.cs         # Custom component base
│   └── Main_Layout.razor                # Layout template
├── _Services/
│   ├── BrowserNewWindow_AppService.cs      # NewWindow_AppService__Interface impl (opens a browser tab)
│   ├── NullNewWindow_AppService.cs         # NewWindow_AppService__Interface no-op (task panes)
│   ├── BrowserExport_AppService.cs        # OfficeExport_AppService__Interface impl (WASM OpenXML)
│   ├── RenderMode_AppService.cs            # Render mode management
│   ├── WatcherEvent_AppService.cs           # Singleton bridge: WinForms Watcher → Blazor Watcher pages
│   └── [Feature]Service.cs              # Blazor services
├── Countries/
│   ├── CountriesNow__Page.razor          # /countries-now — thin routing shell
│   ├── CountriesNow__Component.razor     # grids + export section + terminal log
│   ├── CountriesNow__Component.razor.cs  # partial class; service injections + logic
│   ├── CountrySlides__Page.razor         # /country-slides — thin routing shell
│   ├── CountrySlides__Component.razor    # reveal.js toolbar + slides + jumpbar
│   └── CountrySlides__Component.razor.cs # partial class; slide state + JS interop
├── Watcher/
│   ├── MsWord_Watcher__Page.razor           # /watcher-word — thin routing shell
│   ├── MsWord_Watcher__Component.razor      # open files, save Json, terminal log, events grid
│   ├── MsWord_Watcher__Component.razor.cs   # subscribes to WatcherEvent_AppService.WordStateChanged
│   ├── MsExcel_Watcher__Page.razor          # /watcher-excel — thin routing shell
│   ├── MsExcel_Watcher__Component.razor     # open workbooks, save Json, terminal log
│   ├── MsExcel_Watcher__Component.razor.cs
│   ├── MsPowerPoint_Watcher__Page.razor     # /watcher-powerpoint — thin routing shell
│   ├── MsPowerPoint_Watcher__Component.razor # open presentations, save Json, terminal log
│   └── MsPowerPoint_Watcher__Component.razor.cs
├── Managers/
│   └── [Feature]Manager.razor(.cs)      # Manager components
├── Components/
│   ├── Dialogs/
│   │   └── [Name]Dialog.razor           # Modal dialogs
│   ├── Forms/
│   │   └── [Name]Form.razor             # Form components
│   ├── Lists/
│   │   └── [Name]List.razor             # List/table components
│   └── Viewers/
│       └── [Name]Viewer.razor           # Specialized viewers
├── MainView/
│   └── [Feature]MainView.razor          # Feature main views
├── wwwroot/
│   ├── css/
│   │   ├── app.css
│   │   └── [feature].css                # Feature-specific styles
│   └── js/
│       └── [feature].js                 # Feature-specific scripts
└── bin/, obj/
```

**File Naming:**
- `.razor` for components
- `.razor.cs` for component codebehind
- `Dialog.razor` suffix for modal components
- `Form.razor` suffix for form components
- `List.razor` suffix for list/table components
- `Viewer.razor` suffix for specialized viewers
- `Manager.razor` suffix for manager components
- `MainView.razor` suffix for feature main views

**Component Pattern:**
```
Components/
├── Forms/
│   ├── UserForm.razor                   # Component markup
│   └── UserForm.razor.cs                # Component logic
└── Lists/
    ├── UserList.razor
    └── UserList.razor.cs
```

**_Imports.razor Pattern:**
```razor
@using System.Net.Http
@using System.Net.Http.Json
@using Microsoft.AspNetCore.Components
@using Microsoft.AspNetCore.Components.Forms
@using Microsoft.AspNetCore.Components.Routing
@using Microsoft.AspNetCore.Components.Web
@using Microsoft.JSInterop
@using Radzen
@using Radzen.Blazor
@using JBC.ExploreTheWorld.AL
@using JBC.ExploreTheWorld.AL.BlazorLib
@using JBC.ExploreTheWorld.BL
@using JBC.ExploreTheWorld.CL
```

---

## Web Application (AL.BlazorWebApp)

**Purpose:** ASP.NET Core host for hybrid Blazor app

**Folder Pattern:**
```
AL.BlazorWebApp/
├── ExploreTheWorld.AL.BlazorWebApp.csproj
├── App.razor                            # Root component & HTML layout
├── Program.cs                           # Server initialization
├── Pages/
│   ├── Error.razor                      # Error page
│   └── [FeatureName].razor              # Feature pages
├── Components/
│   ├── Layout/
│   │   ├── MainLayout.razor
│   │   └── NavMenu.razor
│   └── [Feature]Component.razor         # Shared components
├── wwwroot/
│   ├── css/
│   │   ├── app.css
│   │   └── bootstrap.css
│   ├── js/
│   │   └── interop.js
│   └── favicon.png
└── bin/, obj/
```

---

## External REST API Dependency Layer Projects (`DL.{Name}Api`)

Each external API integration is a self-contained project with `_Interfaces/` and `_Rows/` subfolders:

```
DL.{Name}Api/
├── ExploreTheWorld.DL.{Name}Api.csproj
├── _Interfaces/
│   └── {Name}Api_Interface.cs           # Interface for all API operations
├── _Rows/
│   ├── {Entity}_Row.cs                  # JSON response models
│   └── {Nested}_Row.cs                  # Nested model classes
└── {Name}Api__Repo.cs                   # HttpClient implementation
```

Corresponding ._netF project (shared-link compilation):
```
DL.{Name}Api._netF/
└── ExploreTheWorld.DL.{Name}Api._netF.csproj   # Links .cs files from DL.{Name}Api/
```

**Naming rules within these projects:**
- Interface file: `{Name}Api_Interface.cs` — no `I` prefix
- Repo/implementation file: `{Name}Api__Repo.cs` — double underscore
- Row models: `{Entity}_Row.cs` — `_Row` suffix, `[JsonPropertyName]` on every property
- Namespace: `JBC.ExploreTheWorld.DL.{Name}Api`

---

## VBA Projects

```
VBA/
├── Access/
│   ├── ExploreTheWorld.accdb                 # Access database binary (do not edit directly)
│   ├── ExploreTheWorld.laccdb               # Access lock file
│   └── ExploreTheWorld.accdb.src/           # VCS-managed source (MSAccess-VCS 4.0.34) — edit here
│       ├── tables/                              # JSON table schema & data exports
│       ├── forms/                               # Form layout .bas + VBA code-behind .cls
│       │   ├── cns_Country.bas + .cls           # CountriesNowSpace split form
│       │   ├── MsWord.bas + .cls                # Word watcher form
│       │   ├── MsExcel.bas + .cls               # Excel watcher form
│       │   ├── MsPowerPoint.bas + .cls          # PowerPoint watcher form
│       │   └── Main.bas + .cls                  # Main navigation form
│       ├── modules/                             # Standard .bas + class .cls modules
│       │   ├── APP.bas                          # Application helpers (AppendLog, paths)
│       │   ├── ETW__cns_API.bas                 # CountriesNowSpace load/clear
│       │   ├── ETW__cns_MsWord_VBA.cls          # Word export (VB_PredeclaredId=True)
│       │   ├── ETW__cns_MsExcel_VBA.cls         # Excel export
│       │   ├── ETW__cns_MsPowerPoint_VBA.cls    # PowerPoint export
│       │   ├── ETW__FormManager.bas             # OpenWatcher(sType, sFilePath) helper
│       │   └── (JsonConverter, SYS__Guid, mFile, ...)
│       ├── queries/
│       └── relations/
├── PowerPoint/
│   ├── ExploreTheWorld.pptm                 # Macro-enabled presentation (do not edit directly)
│   └── ExploreTheWorld.pptm.src/            # VBA source — edit here
│       ├── ThisPresentation.cls             # IRibbonExtensibility; delegates ribbon callbacks
│       ├── RibbonCallbacks.bas              # GetRibbonXml() returning ETW tab XML
│       ├── ETW__PowerPoint_CountriesNow_Form.frm + .frx  # Export form
│       └── ETW__PowerPoint_Watcher_Form.frm + .frx       # Watcher form
├── Word/
│   ├── ExploreTheWorld.dotm                 # Macro-enabled template (do not edit directly)
│   └── ExploreTheWorld.dotm.src/            # VBA source — edit here
│       ├── ThisDocument.cls                 # IRibbonExtensibility
│       ├── RibbonCallbacks.bas
│       ├── ETW__Word_CountriesNow_Form.frm + .frx
│       └── ETW__Word_Watcher_Form.frm + .frx
└── Excel/
    ├── ExploreTheWorld.xlsm                 # Macro-enabled workbook (do not edit directly)
    └── ExploreTheWorld.xlsm.src/            # VBA source — edit here
        ├── ThisWorkbook.cls                 # IRibbonExtensibility
        ├── RibbonCallbacks.bas
        ├── ETW__Excel_CountriesNow_Form.frm + .frx
        └── ETW__Excel_Watcher_Form.frm + .frx
```

---

## WinForms Application Projects

### AL.WinFormApp (net10.0 — BlazorWebView)

```
AL.WinFormApp/
├── ExploreTheWorld.AL.WinFormApp.csproj   # SDK=Microsoft.NET.Sdk.Razor, net10.0-windows
├── Program.cs                             # STAThread, DI setup (ShowSidebar=true); launches ExploreTheWorld_Form
├── _Forms/
│   ├── ExploreTheWorld_Form.cs + .Designer.cs # Single BlazorWebView host: all pages + Word COM watcher
│   ├── CountriesNowSpace_Form.cs + .Designer.cs + .resx      # Traditional WinForms form (non-BlazorWebView)
│   ├── CountriesNowSpace_UserControl.cs + .Designer.cs        # Canonical UC; shared to ._netF + VstoAddIns
│   └── _Watcher/                             # Canonical source (shared to ._netF and VstoAddIns)
│       ├── WatcherInteropMethod_Enum.cs      # DotNetInterop | LateBinding | NetOffice
│       ├── MsOfficeEvent_Record.cs           # Mutable event row (Name, Category, Log checkbox)
│       ├── MsOfficeEvents_Repo.cs            # Hard-coded Office event catalogue
│       ├── MsOfficeJsonWriter_Helper.cs      # Writes open document to JSON
│       ├── WatcherComHelper.cs               # GetActiveCom() — abstracts GetActiveObject across frameworks
│       ├── MsWord_Watcher_Form.cs + .Designer.cs      # Traditional Word watcher (DotNetInterop/LateBinding/NetOffice)
│       ├── MsExcel_Watcher_Form.cs + .Designer.cs     # Traditional Excel watcher
│       ├── MsPowerPoint_Watcher_Form.cs + .Designer.cs # Traditional PowerPoint watcher
│       ├── MsWord_Watcher_UserControl.cs + .Designer.cs     # UC shared to VSTO addins via <Compile Include>
│       ├── MsExcel_Watcher_UserControl.cs + .Designer.cs
│       └── MsPowerPoint_Watcher_UserControl.cs + .Designer.cs
├── _Services/
│   ├── OfficeExport_AppService.cs             # OfficeExport_AppService__Interface for WinForms; delegates to BL MsOfficeExportManager
│   ├── MsOfficeExportRepoFactory.cs       # MsOfficeExportRepoFactory__Interface impl (news OpenXml/NetOffice/Interop DL repos)
│   └── WinFormNewWindow_AppService.cs         # NewWindow_AppService__Interface: opens new ExploreTheWorld_Form
└── wwwroot/
    └── index.html                             # Blazor WebView host page
```

**Single-form design (ExploreTheWorld_Form):** The single `ExploreTheWorld_Form` replaces the previous `Main_Form` + separate per-page WebView forms. It hosts a fill-docked `BlazorWebView` with the full `Routes` component and Radzen sidebar enabled, so users navigate between Countries Now, Country Slides, and Word/Excel/PowerPoint Watcher pages via the sidebar. Word COM connection is established on demand by clicking **Connect** on the Word Watcher page; Excel and PowerPoint COM connections are not yet implemented in standalone mode.

### AL.WinFormApp._netF (.NET Framework 4.8.1 — Traditional WinForms)

```
AL.WinFormApp._netF/
├── ExploreTheWorld.AL.WinFormApp._netF.csproj   # SDK=Microsoft.NET.Sdk, net481
├── Program.cs                                    # EnableVisualStyles, direct service instantiation; launches Main_Form
├── App.config                                    # .NET Framework startup config
└── _Forms/
    ├── Main_Form.cs + .Designer.cs + .resx          # Own file; not linked from AL.WinFormApp
    └── CountriesNowSpace_Form.resx                  # Own resx; .cs/.Designer.cs linked from AL.WinFormApp
    # CountriesNowSpace_UserControl, _Watcher helper + UC .cs files all linked from AL.WinFormApp
    # No WebView variants (BlazorWebView requires net10.0)
```

---

## Office Add-in Projects

## Office Add-in Projects

### AL.MsOffice{Word|Excel|PowerPoint}VstoAddIn (net10.0 — COM add-in via comhost)

Each host has a matching project folder. The pattern is identical across all three hosts.

```
AL.MsOffice{Host}VstoAddIn/
├── ExploreTheWorld.AL.MsOffice{Host}VstoAddIn.csproj  # SDK=WindowsDesktop, net10.0-windows, EnableComHosting=true
├── Addin.cs                                            # IDTExtensibility2 + IRibbonExtensibility
├── RuntimeManifestFileProvider.cs                      # Reads *.staticwebassets.runtime.json; fixes "Loading..." in COM host
├── RibbonUI.xml                                        # Embedded ribbon: "ETW (VSTO)" tab — 3 groups, 4 buttons
└── _Forms/                                             # Linked files from AL.WinFormApp
    ├── CountriesNowSpace_UserControl.cs + .Designer.cs   # (linked)
    └── _Watcher/
        ├── MsOfficeEvent_Record.cs                         # (linked)
        ├── MsOfficeEvents_Repo.cs                          # (linked)
        ├── MsOfficeJsonWriter_Helper.cs                    # (linked)
        ├── WatcherComHelper.cs                             # (linked)
        ├── Ms{Host}_Watcher_Form.cs + .Designer.cs         # (linked) floating form
        └── Ms{Host}_Watcher_UserControl.cs + .Designer.cs  # (linked) task pane UC
```

**Ribbon tab "ETW (VSTO)" — 3 groups (Countries API, Watcher, Export):**
- **Countries API › Countries Form (Blazor)** — opens the floating `CountriesNowSpace_WebView_Form` (no file/type/library selection — uses the active document)
- **Watcher › Watcher Form (Blazor)** — opens the floating `Ms{Host}_Watcher_WebView_Form`
- **Export › Save as JSON** — C# NetOffice writer via a Save dialog
- **Export › Save as JSON (Direct)** — runs the `MSO_Ms{Host}.WriteActive…` VBA macro (the "Direct" writer)

**Build target `RegisterComAddin64`** (Debug only, in `.csproj`): writes all required `HKCU` registry keys after build (Addin CLSID → `comhost.dll`, `HKCU\Software\Microsoft\Office\{Host}\Addins\{ProgId}`).

---

### AL.MsOffice{Word|Excel|PowerPoint}VstoAddIn._netF (net481 — COMAddin base)

```
AL.MsOffice{Host}VstoAddIn._netF/
├── ExploreTheWorld.AL.MsOffice{Host}VstoAddIn._netF.csproj  # SDK=Microsoft.NET.Sdk, net481; MSOFFICE_ADDIN symbol
├── Addin.cs                                                   # Inherits COMAddin; [CustomPane] ×2 (Watcher + CountriesNow)
├── RibbonUI.xml                                               # Embedded ribbon: "ETW (VSTO._netF)" tab — 3 groups, 6 buttons
├── Ms{Host}_Watcher_UserControl.cs + .Designer.cs             # Canonical own file
├── CountriesNow_TaskPane_UserControl.cs + .Designer.cs        # Wrapper UC (parameterless ctor, creates manager chain)
    # CountriesNowSpace_Form + CountriesNowSpace_UserControl linked from AL.WinFormApp._netF
    # _Watcher helpers (JsonWriters, Event_Record, Events_Repo, ComHelper) linked from AL.WinFormApp
```

**Key files:**
- `Addin.cs` — inherits `COMAddin`; `[CustomPane]×2` (Watcher at index 0, Countries API at index 1); 3-group ribbon (Countries API, Watcher, Export); `CustomUI_OnLoad` logs pane count for diagnostics; toggle buttons sync via `RibbonUI.InvalidateControl` for both controls
- `CountriesNow_TaskPane_UserControl.cs` — wrapper with parameterless ctor so `ICTPFactory.CreateCTP` can instantiate it; manually creates `ExploreTheWorldDbContextFactory → CountriesNowSpaceApiManager__Repo → CountriesNowSpaceManager__Service`; embeds `CountriesNowSpace_UserControl` dock-filled
- Watcher UserControl file remains a local own file (canonical source for ._netF pattern)

**Build target `RegisterComAddin`** (Debug only, in `.csproj`): registers Addin CLSID (via `mscoree.dll`) and **all three** UserControl CLSIDs (Watcher + Countries API pane).

---

### AL.MsOffice{Word|Excel|PowerPoint}BlazorWebAddIn + .Client (Office.js Web Add-in)

Each host add-in is a server project plus a separate WASM client project:

```
AL.MsOffice{Host}BlazorWebAddIn/
├── ExploreTheWorld.AL.MsOffice{Host}BlazorWebAddIn.csproj  # Server host (net10.0, SDK=Microsoft.NET.Sdk.Web)
├── package.json                                             # Office add-in local debug tooling
├── Assets/
│   └── manifest.local.xml                                   # Local manifest for debugging
└── AL.MsOffice{Host}BlazorWebAddIn.Client/
    ├── ExploreTheWorld.AL.MsOffice{Host}BlazorWebAddIn.Client.csproj  # WASM client (net10.0)
    ├── Program.cs
    ├── Pages/
    │   ├── Home.razor + .cs + .js
    │   ├── DocumentInfo.razor + .cs
    │   ├── CountriesNowSpace.razor + .cs     # CNS API + "Save as JSON" download button
    │   └── Watcher.razor + .cs + .js         # Office.js event monitor
    └── Services/
```

**CountriesNowSpace.razor:** calls countriesnow.space API via `CountriesNowSpaceApi__Repo` (injected `HttpClient`); displays countries; **Save as JSON** button triggers JS interop `window.downloadJson(filename, json)` browser download. No Export Type, Export Library, or file path selection.

**Watcher.razor:** monitors Office.js events for the active document; logs event name + timestamp; `[JSInvokable] OnEventLogged` callback from JS.

**Manifest ribbon (manifest.local.xml):** "ETW (Web)" tab — 3 groups, one button each:
- **Countries API › Countries Pane (Blazor)** — opens the task pane to `/countries-now`
- **Watcher › Watcher Pane (Blazor)** — opens the task pane to `/events`
- **Export › Save as JSON (Blazor)** — opens the task pane to `/save-as-json`

**Sideload location:** manifests copied to `_sideload/{Host}/Web/` for manual sideloading.

---

**Purpose:** Client-side Blazor WebAssembly app

**Folder Pattern:**
```
AL.BlazorWebApp.Client/
├── ExploreTheWorld.AL.BlazorWebApp.Client.csproj
├── App.razor                            # Root component
├── Program.cs                           # Client initialization
├── Pages/
│   ├── Index.razor                      # Home page
│   └── [FeatureName].razor              # Feature pages
├── Shared/
│   └── NavMenu.razor                    # Navigation component
├── _Imports.razor                       # Global usings
├── wwwroot/
│   ├── css/
│   │   └── app.css
│   ├── js/
│   │   └── interop.js
│   └── index.html
└── bin/, obj/
```

---

## Standalone WebAssembly PWA (AL.BlazorWebApp.ClientOnly)

**Purpose:** Self-contained WASM app served as static files; no ASP.NET server needed at runtime.

**Folder Pattern:**
```
AL.BlazorWebApp.ClientOnly/
├── ExploreTheWorld.AL.BlazorWebApp.ClientOnly.csproj  # SDK=BlazorWebAssembly
├── Program.cs                           # WebAssemblyHostBuilder; mounts AL.BlazorLib.Routes
├── _Imports.razor                       # Global usings
├── wwwroot/
│   ├── index.html                       # Entry page (Radzen CSS, BlazorLib CSS; JS interop is module-loaded, no global scripts)
│   ├── service-worker.js                # Dev no-op
│   ├── service-worker.published.js      # Prod offline cache
│   ├── manifest.webmanifest             # PWA manifest
│   └── favicon.png
└── bin/, obj/
```

**Notes:**
- Uses `AL.BlazorLib/Routes.razor` (shared) — no own `Routes.razor`
- Registers `CountriesNowSpaceApiManager__WasmNoCache__Repo` (SQLite unavailable in WASM)
- Registers `BrowserExport_AppService` for in-browser OpenXML export
- Static assets from `AL.BlazorLib` served via `_content/JBC.ExploreTheWorld.AL.BlazorLib/`

---

## File Naming Conventions by Type

### C# Classes

| Type | Pattern | Example |
|------|---------|---------|
| Entity Model | `{Name}_Model.cs` | `User_Model.cs` |
| Repository Interface | `{Name}__Repo__Interface.cs` | `UserRepository__Repo__Interface.cs` |
| Repository Implementation | `{Name}__Repo.cs` | `UserRepository__Repo.cs` |
| Manager | `{Name}Manager.cs` | `UserManager.cs` |
| Validator | `{Name}Validator.cs` | `UserValidator.cs` |
| Service | `{Name}Service.cs` | `AuthenticationService.cs` |
| Logger/Extension | `{Name}_Extensions.cs` | `Enum_Extensions.cs` |
| Base Class | `Base__{Name}.cs` | `Base__RadzenComponent.cs` |
| Configuration | `{Name}Config.cs` | `ServiceConfig.cs` |

### Razor Components

| Type | Pattern | Example |
|------|---------|---------|
| Page (routing shell) | `{Name}__Page.razor` | `CountriesNow__Page.razor` |
| Reusable Component | `{Name}__Component.razor` | `CountriesNow__Component.razor` |
| Component Codebehind | `{Name}__Component.razor.cs` | `CountriesNow__Component.razor.cs` |
| Layout | `{Name}__Layout.razor` | `Main_Layout.razor` |
| Dialog/Modal | `{Name}__Dialog.razor` | `ConfirmDelete__Dialog.razor` |

### Markup Files

| Type | Pattern | Example |
|------|---------|---------|
| Global Imports | `_Imports.razor` | `_Imports.razor` |
| App Root | `App.razor` | `App.razor` |
| Error Page | `Error.razor` | `Error.razor` |
| Layout | `*.razor` | `MainLayout.razor` |
| Stylesheet | `*.css` | `app.css` |
| Script | `*.js` | `interop.js` |

---

## Namespace Organization

Namespaces follow the folder structure:

```
Folder: DL/Domain/Users/UserRepository.cs
Namespace: JBC.ExploreTheWorld.DL.Users

Folder: AL.BlazorLib/Components/Forms/UserForm.razor.cs
Namespace: JBC.ExploreTheWorld.AL.BlazorLib.Components.Forms

Folder: BL/Domain/Users/UserManager.cs
Namespace: JBC.ExploreTheWorld.BL.Users

Folder: BL/_Services/UserManager__Service.cs
Namespace: JBC.ExploreTheWorld.BL          (leading-underscore folder "_Services" is dropped)
```

---

## Best Practices

### 1. One Type Per File
Each `.cs` file contains exactly one type — class, enum, interface, record, or struct — whether public or internal, and the file is named after the type:
- `Country.cs` → `public class Country`
- `TriState_Enum.cs` → `public enum TriState_Enum` (companion extension classes like `TriStateExtensions` get their own file)
- Code-behind files (`Component.razor.cs`) contain only that component's partial class; supporting rows/records move to their own files
- Third-party content (`oqtane.framework/`, `node_modules/`) is exempt

### 2. Logical Grouping
Organize related files in subfolders:
```
DL/Domain/Users/
  ├── User_Model.cs
  ├── IUserRepository.cs
  └── UserRepository.cs
```

### 3. Domain-First Organization
Group by business domain before type:
```
DL/Domain/
  ├── Users/
  ├── Content/
  └── Viewers/
```
NOT by type:
```
DL/Models/
DL/Repositories/
```

### 4. Consistent Casing
- Folder names: PascalCase (Users, Components, Domain)
- File names: PascalCase (.cs files), kebab-case (wwwroot assets)
- Class names: PascalCase
- Method names: PascalCase
- Property names: PascalCase
- Variable names: camelCase

### 5. Related Files Together
Place codebehind and related files next to their parent:
```
Components/Forms/
  ├── UserForm.razor
  └── UserForm.razor.cs
```

### 6. Shared Resources in wwwroot
- Shared stylesheets in `wwwroot/css/`
- Feature-specific styles in feature folders
- Scripts in `wwwroot/js/`
- Images and other assets organized similarly

### 7. Clean Separation
- No circular dependencies
- Clear layer boundaries
- Interfaces to abstract dependencies
