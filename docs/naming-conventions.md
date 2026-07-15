# Naming Conventions Document

## Overview

Consistent naming conventions across the codebase enhance readability, maintainability, and reduce cognitive load when navigating projects. This document defines naming conventions for files, types, members, and identifiers in the ExploreTheWorld solution.

---

## General Principles

1. **Use Meaningful Names** - Names should clearly indicate purpose/responsibility
2. **Be Consistent** - Apply patterns uniformly across the codebase
3. **Use English** - All names in English for international accessibility
4. **Avoid Abbreviations** - Spell out terms unless they're well-known domain terms
5. **Context Matters** - Shorter names acceptable in limited scope, longer names for broader scope

---

## File Naming Conventions

### C# Source Files

#### Pattern by Type

| Entity Type | File Pattern | Example | Location |
|-------------|-------------|---------|----------|
| Domain Model | `{EntityName}_Model.cs` | `User_Model.cs` | `Domain/{Domain}/` |
| Repository Interface | `{EntityName}__Repo__Interface.cs` | `UserRepository__Repo__Interface.cs` | `Repositories/` |
| Repository Implementation | `{EntityName}__Repo.cs` | `UserRepository__Repo.cs` | `Repositories/` |
| Business Manager | `{DomainName}Manager.cs` | `UserManager.cs` | `Managers/` |
| Validator Class | `{EntityName}Validator.cs` | `UserValidator.cs` | `Managers/` |
| BL Service | `{Name}__Service.cs` | `CountriesNowSpaceManager__Service.cs` | `BL/_Services/` |
| AL App Service | `{Name}_AppService.cs` | `Layout_AppService.cs` | `AL*/_Services/` |
| AL App Service Interface | `{Name}_AppService__Interface.cs` | `NewWindow_AppService__Interface.cs` | `AL/_Interfaces/` |
| Extension Methods | `{TargetType}_Extensions.cs` | `String_Extensions.cs` | `_Utilities/` or `Enum_Extensions.cs` in root |
| Base/Abstract Class | `Base__{ClassName}.cs` | `Base__RadzenComponent.cs` | Feature folder |
| Utility Class | `{Purpose}Utilities.cs` | `ValidationUtilities.cs` | `_Utilities/` |
| Configuration | `{Feature}Config.cs` | `DatabaseConfig.cs` | Project root |
| Constants | `{Domain}Constants.cs` | `UserConstants.cs` | Project root or domain folder |
| Enum | `{Name}Enum.cs` | `UserRoleEnum.cs` | Domain folder |
| Helper | `{Name}Helper.cs` | `DateHelper.cs` | `_Utilities/` |
| Factory | `{Type}Factory.cs` | `RepositoryFactory.cs` | `Factories/` |
| Options/Settings | `{Feature}Options.cs` | `DatabaseOptions.cs` | `Configuration/` |

#### CL (Common Layer) Special Files

| Type | Pattern | Example |
|------|---------|---------|
| Entity Base | `_Row.cs` | Located in `_Row/` subfolder |
| Row Audit Log | `RowLog_Row.cs` | Located in `_Row/` subfolder |
| Column Audit Log | `ColumnLog_Row.cs` | Located in `_Row/` subfolder |
| Enum Extensions | `Enum_Extensions.cs` | Located in `Enum/` subfolder |
| Logger Interface | `ILogger2.cs` | Located in `_Services/` subfolder |
| Tree Structure | `TreeNode.cs` | Located in `TreeStructure/` subfolder |

### Razor Components

Double-underscore suffixes identify Razor file roles at a glance:

| Component Type | File Pattern | Example | Location |
|--|--|--|--|
| Page/Route Component | `{Name}__Page.razor` | `CountriesNow__Page.razor` | Domain folder |
| Reusable Component | `{Name}__Component.razor` | `CountriesNow__Component.razor` | Domain folder |
| Layout Component | `{Name}__Layout.razor` | `Main_Layout.razor` | `_Shared/` |
| Dialog/Modal | `{Name}__Dialog.razor` | `ConfirmDelete__Dialog.razor` | Domain folder |
| Codebehind | `{Name}__Component.razor.cs` | `CountriesNow__Component.razor.cs` | Same folder as .razor |
| Global Imports | `_Imports.razor` | Fixed | Project root |

**Page pattern:** A `__Page` is a thin routing shell — `@page` directive and one `<{Name}__Component />` call. No markup beyond that. It does **not** set its own `<PageTitle>`: the title bar is supplied centrally by `Main_Layout`, which renders `Layout_AppService.AppTitle` (each host's short project name, e.g. `ETW BlazorWebApp`) as the single `<PageTitle>`. See the Branding section in [AGENTS.md](../AGENTS.md).

**Component pattern:** A `__Component` contains all markup and logic. The `.razor` file has `@inherits Base__RadzenComponent` and the standard Radzen root-div wrapper. The `.razor.cs` is a `partial class {Name}__Component : Base__RadzenComponent` with injected services and private fields/methods.

```razor
@* {Name}__Page.razor *@
@page "/route"
@namespace JBC.ExploreTheWorld.AL.BlazorLib.{Domain}

@* Title bar comes from Main_Layout (Layout_AppService.AppTitle) — no per-page <PageTitle>. *@
<{Name}__Component Style="height:100%" />
```

```razor
@* {Name}__Component.razor *@
@inherits Base__RadzenComponent
@namespace JBC.ExploreTheWorld.AL.BlazorLib.{Domain}

@if (!Visible) return;

<div id="@GetId()" class="@GetCssClass()" style="@GetHiddenStyle()" @ref="@Element" @attributes="Attributes">
    <!-- markup here -->
</div>
```

```csharp
// {Name}__Component.razor.cs
public partial class {Name}__Component : Base__RadzenComponent
{
    [Inject] protected Layout_AppService Layout_AppService { get; set; } = default!;  // injected service members use _AppService (see Service Naming)
    // private fields and methods (no protected needed — same partial class)
}
```

### Configuration & Metadata

| Type | Pattern | Example |
|------|---------|---------|
| Solution File | `JBC.{ProjectName}.sln` | `JBC.ExploreTheWorld.sln` |
| Project File | `ExploreTheWorld.{Layer}[.Specialty].csproj` | `ExploreTheWorld.AL.BlazorLib.csproj` |
| Web Assets | `{name}.{ext}` | `app.css`, `interop.js` (kebab-case for URLs) |
| Configuration JSON | `appsettings.{Environment}.json` | `appsettings.json`, `appsettings.Production.json` |

---

## Namespace Conventions

### Namespace Structure Pattern

```
JBC.ExploreTheWorld[.{Layer}][.{Specialty}][.{Domain}][.{SubDomain}]
```

### Layer Namespaces

| Layer | Namespace Pattern | Example |
|-------|-------------------|---------|
| CL | `JBC.ExploreTheWorld.CL` | `JBC.ExploreTheWorld.CL.Enum` |
| DL | `JBC.ExploreTheWorld.DL` | `JBC.ExploreTheWorld.DL.Users` |
| BL | `JBC.ExploreTheWorld.BL` | `JBC.ExploreTheWorld.BL` (services live in `BL/_Services/`; the `_`-folder is dropped) |
| AL | `JBC.ExploreTheWorld.AL` | `JBC.ExploreTheWorld.AL.Services` |
| AL.BlazorLib | `JBC.ExploreTheWorld.AL.BlazorLib` | `JBC.ExploreTheWorld.AL.BlazorLib.Components.Forms` |
| AL.BlazorWebApp | `JBC.ExploreTheWorld.AL.BlazorWebApp` | `JBC.ExploreTheWorld.AL.BlazorWebApp.Pages` |
| AL.BlazorWebApp.Client | `JBC.ExploreTheWorld.AL.BlazorWebApp.Client` | `JBC.ExploreTheWorld.AL.BlazorWebApp.Client.Pages` |

### Namespace by Folder Location

Namespace should mirror the folder structure:

```
File: DL/Domain/Users/UserRepository.cs
Namespace: JBC.ExploreTheWorld.DL.Users

File: AL.BlazorLib/Components/Forms/UserForm.razor.cs
Namespace: JBC.ExploreTheWorld.AL.BlazorLib.Components.Forms

File: DL.MsSystemNet/Http/RestClient.cs
Namespace: JBC.ExploreTheWorld.DL.MsSystemNet.Http
```

#### Underscore-prefixed folders are excluded from the namespace

**A folder whose name begins with `_` is a grouping/organizational folder only — it does NOT contribute a segment to the namespace.** Files inside it belong to the namespace of the nearest non-underscore ancestor folder. This applies to `_Services/`, `_Export/`, `_Watcher/`, `_Forms/`, `_Shared/`, `_Utilities/`, `_Row/`, etc.

```
File: BL/_Services/CountriesNowSpaceManager__Service.cs
Namespace: JBC.ExploreTheWorld.BL            ← "_Services" is dropped

File: AL.WinFormsLib/_Services/OfficeExport_AppService.cs
Namespace: JBC.ExploreTheWorld.AL.WinFormsLib ← "_Services" is dropped

File: AL.WinFormsLib/_Export/ExportMenuHelper.cs
Namespace: JBC.ExploreTheWorld.AL.WinFormsLib ← "_Export" is dropped
```

#### Group by type, not by feature — and only when there are 2+

Within a project, group interchangeable members **by type** in an underscore folder: repositories **and their repo interfaces** in `_Repos/`, factories in `_Factories/`, EF entities in `_Entities/`, DTO rows in `_Rows/`, EF field-name constants in `_Fields/`. Do **not** introduce per-feature grouping folders (no `Managers/`, no `CountriesNowSpaceApiManager/`, no `MsOfficeDocumentManager/`) — they add a redundant namespace segment.

**An underscore folder is only warranted when it groups 2+ members of that type.** A project with a single repo (or single factory) places it at the project root, not in a `_Repos/` (or `_Factories/`) folder. Example: each `*Db_Impl` has one repo and one factory → both at the project root; `IndexedDb_Impl` has three repos → all in `_Repos/`.

### Domain-Organized Namespaces

When organizing by business domain:

```
DL/Domain/Users/
  └── UserRepository.cs
     Namespace: JBC.ExploreTheWorld.DL.Users

DL/Domain/Content/
  └── ContentRepository.cs
     Namespace: JBC.ExploreTheWorld.DL.Content

DL/Domain/Viewers/
  └── ViewerRepository.cs
     Namespace: JBC.ExploreTheWorld.DL.Viewers
```

---

## Class & Member Naming

### Class Naming

| Type | Pattern | Example | Notes |
|------|---------|---------|-------|
| Regular Class | `PascalCase` | `UserManager`, `OrderService` | Standard C# convention |
| Entity/Model | `{Name}_Model` | `User_Model`, `Product_Model` | Suffix indicates entity |
| Repository (DL) | `{Entity}__Repo` | `UserRepository__Repo` | Double-underscore suffix; DL data-access implementations |
| Service (BL) | `{Name}__Service` | `CountriesNowSpaceManager__Service`, `DbProviderSwitcher__Service` | BL business-logic orchestrator; lives in `BL/_Services/`. Injects DL repo interfaces; **has no interface of its own** (mock via the injected DL interfaces) |
| App Service (AL) | `{Name}_AppService` | `Layout_AppService`, `BrowserExport_AppService`, `MauiWatcher_AppService` | AL application/platform service (host plumbing). Single-underscore `_AppService` keeps it visually distinct from a BL `__Service`. Concrete by default; add an `_AppService__Interface` only when hosts swap implementations |
| Interface | `{Name}__Repo__Interface` / `{Name}_AppService__Interface` | `UserRepository__Repo__Interface`, `NewWindow_AppService__Interface` | `_Interface` suffix, no `I` prefix. BL services are concrete (no interface); AL app-service seams and host-implemented seams like the export-repo factory (a **DL** contract, `DL.MsOfficeApi.MsOfficeExportRepoFactory__Interface`) do define interfaces |
| Factory | `{Name}__Factory` | `ExploreTheWorldDbContext__SqliteDb__Factory` | `__Factory` suffix. Grouped in a `_Factories/` folder when 2+ exist; a lone factory sits at the project root |
| Base Class | `Base__{Name}` | `Base__RadzenComponent` | Clear base class indicator |
| Abstract Class | `Abstract_{Name}` | `Abstract_Service` | Less common, use base__ pattern |
| Enum | `{Name}Enum` | `UserRoleEnum`, `StatusEnum` | Optional suffix for clarity |
| Static Utility | `{Purpose}Utilities` | `ValidationUtilities` | Indicates static-only class |

### Property Naming

| Type | Pattern | Example | Notes |
|------|---------|---------|-------|
| Public Property | `PascalCase` | `UserId`, `FirstName`, `IsActive` | Standard C# convention |
| Auto-Property | `PascalCase` | `public string Name { get; set; }` | getter/setter on same line for simple props |
| Backing Field | `_camelCase` | `private string _name;` | Private field with underscore prefix |
| Constant | `UPPER_CASE` or `PascalCase` | `MaxRetries = 3` or `MAX_RETRIES` | Use PascalCase for class constants, UPPER for static const |

### Method Naming

| Type | Pattern | Example | Notes |
|------|---------|---------|-------|
| Public Method | `PascalCase` | `GetUser()`, `CalculateTotal()` | Standard C# convention |
| Async Method | `{Action}Async` | `GetUserAsync()`, `SaveAsync()` | Always suffix with Async |
| Event Handler | `{Source}_{Event}` | `Button_Click()`, `Form_Submit()` | Clear source and event |
| Private Method | `PascalCase` | `ValidateInput()` | Same as public, differentiated by access modifier |
| Query Method | `Get{Entity}` / `Get{Entities}` | `GetUser()`, `GetAllUsers()` | Clear query intent |
| Command Method | `{Verb}{Entity}` | `CreateUser()`, `DeleteUser()`, `UpdateUser()` | Clear action intent |
| Predicate Method | `Is{Condition}` / `Has{Feature}` | `IsActive()`, `HasPermission()` | Returns boolean |
| Converter Method | `To{Type}` | `ToString()`, `ToViewModel()` | Conversion method |

### Variable Naming

| Type | Pattern | Example | Notes |
|------|---------|---------|-------|
| Local Variable | `camelCase` | `userName`, `totalAmount`, `itemCount` | Standard C# convention |
| Parameter | `camelCase` | `public void SetUser(string userName)` | Standard C# convention |
| Loop Variable | `camelCase` or single letter | `for (int i = 0; i < count; i++)` | Single letter acceptable for simple loops |
| Lambda Parameter | `camelCase` | `users.Where(u => u.IsActive)` | Abbreviated names acceptable |

---

## Domain-Specific Naming

### Common Domain Names

Use these domain names consistently across the codebase:

| Domain | Examples |
|--------|----------|
| Users | User, UserManager, UserRepository, UserValidator |
| Authentication | Auth, AuthenticationService, LoginForm |
| Authorization | Permission, Role, RoleValidator |
| Content | Content, Document, Media |
| Viewers | Viewer, ViewerManager, ViewerComponent |
| Settings | Setting, Configuration, Config |
| Logging | Log, Audit, RowLog, ColumnLog |

### Entity Model Naming

Entity models (inheriting from `_Row`) use the `_Model` suffix:

```csharp
public class User_Model : _Row
{
    public string Email { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
}

public class Content_Model : _Row
{
    public string Title { get; set; }
    public string Body { get; set; }
}
```

### Manager Class Naming

Manager classes handle business logic for a domain:

```csharp
namespace JBC.ExploreTheWorld.BL.Users
{
    public class UserManager
    {
        public async Task<User_Model> CreateUserAsync(CreateUserRequest request) { }
        public async Task<User_Model> GetUserAsync(Guid userId) { }
        public async Task UpdateUserAsync(User_Model user) { }
    }
}
```

### Repository Naming

Repository classes handle data access:

```csharp
public interface UserRepository__Repo__Interface
{
    Task<User_Model> GetAsync(Guid id);
    Task<IEnumerable<User_Model>> GetAllAsync();
    Task CreateAsync(User_Model entity);
    Task UpdateAsync(User_Model entity);
    Task DeleteAsync(Guid id);
}

public class UserRepository__Repo : UserRepository__Repo__Interface
{
    // Implementation
}
```

---

## Component Naming (Blazor)

### Dialog Components

```csharp
// File: ConfirmDeleteDialog.razor
@namespace JBC.ExploreTheWorld.AL.BlazorLib.Components.Dialogs

<RadzenDialog Title="Confirm Delete">
    <p>Are you sure?</p>
    <RadzenButton Text="Delete" Click="OnConfirm" />
</RadzenDialog>

@code {
    [Parameter] public EventCallback<bool> OnConfirm { get; set; }
}
```

Naming:
- File: `{Action}Dialog.razor` (e.g., `ConfirmDeleteDialog.razor`)
- Component: `{Action}Dialog` class (auto-generated from filename)

### Form Components

```csharp
// File: UserForm.razor
@namespace JBC.ExploreTheWorld.AL.BlazorLib.Components.Forms

<EditForm Model="User" OnValidSubmit="HandleSubmit">
    <InputText @bind-Value="User.Name" />
    <button type="submit">Save</button>
</EditForm>

@code {
    [Parameter] public User_Model User { get; set; }
    [Parameter] public EventCallback<User_Model> OnSubmit { get; set; }
}
```

Naming:
- File: `{EntityName}Form.razor` (e.g., `UserForm.razor`)
- Codebehind: `{EntityName}Form.razor.cs`
- Parameter methods: `On{Action}` (e.g., `OnSubmit`, `OnCancel`)

### List Components

```csharp
// File: UserList.razor
@namespace JBC.ExploreTheWorld.AL.BlazorLib.Components.Lists

<RadzenDataGrid Data="Users">
    <Columns>
        <RadzenDataGridColumn Property="Name" Title="Name" />
    </Columns>
</RadzenDataGrid>
```

Naming:
- File: `{EntityName}List.razor`
- Data property: `{Entities}` (plural, e.g., `Users`)

### Viewer Components

```csharp
// File: DocumentViewer.razor
@namespace JBC.ExploreTheWorld.AL.BlazorLib.Components.Viewers

<div class="viewer">
    @((MarkupString)Content.Body)
</div>
```

Naming:
- File: `{EntityName}Viewer.razor`
- Used for read-only display of complex entities

---

## Service Naming

**Two distinct service tiers, deliberately spelled differently so they never blur together at a glance:**

| Tier | Suffix | Example | Interface |
|------|--------|---------|-----------|
| **BL** business-logic orchestrator | `__Service` (double underscore) | `CountriesNowSpaceManager__Service` | none (concrete; mock its injected DL repo interfaces) |
| **AL** application/platform service | `_AppService` (single underscore) | `Layout_AppService`, `BrowserExport_AppService` | optional `_AppService__Interface` when a host swaps implementations |

The `App` prefix on the `Service` word marks a service as belonging to the **application/host layer** (AL) — platform-specific plumbing (Blazor render-mode/layout state, new-window and export abstractions, watcher plumbing) — so it stands out from a BL `__Service`. **Every `Service`-named type in an `AL*` project uses `_AppService`; never a bare `_Service` or `Service` suffix.**

**Type names vs. injected-member names.** Third-party service *types* keep their own names — we cannot (and do not) rename `Radzen.DialogService`, `IServiceCollection`, `NavigationManager`, etc. But the **fields/properties/variables that hold an injected service** — ours *and* third-party — use the `_AppService` suffix so every service member reads consistently at the injection site (see the next subsection).

### AL app services (stateful, injected per component/host)

```csharp
// Concrete AL app service — no interface needed
public class RenderMode_AppService
{
    public void SetRenderMode(IComponentRenderMode mode) { }
    public IComponentRenderMode GetRenderMode() { }
}

// Register: builder.Services.AddScoped<RenderMode_AppService>();
```

Naming:
- Suffix: `_AppService` (e.g., `RenderMode_AppService`)
- Action methods: `{Verb}{Noun}` (e.g., `SetRenderMode`, `GetRenderMode`)

### AL app services with a host-swappable implementation

When several hosts supply platform-specific implementations of the same seam (e.g. "open a new window"), define an interface in the core `ExploreTheWorld.AL` project and implement it per host:

```csharp
// Contract in core AL: ExploreTheWorld.AL/_Interfaces/
public interface NewWindow_AppService__Interface
{
    Task OpenNewWindowAsync(string url);
}

// Per-host implementation (AL.BlazorLib, AL.WinFormsLib, AL.MauiLib, …)
public class BrowserNewWindow_AppService : NewWindow_AppService__Interface
{
    // Implementation
}

// Register: builder.Services.AddScoped<NewWindow_AppService__Interface, BrowserNewWindow_AppService>();
```

### Injected service member and variable names

Any **field, property, or local variable that holds an injected service** uses the `_AppService` suffix on its own role-based base name — regardless of whether the service *type* is one of ours or a third-party type. Keep the member's descriptive base (`Layout`, `Export`, `Watcher`, `Dialog`, …) and end it in `_AppService`; do **not** copy the full type name.

```csharp
public partial class CountriesNow__Component : Base__RadzenComponent
{
    // AL app-service members — base name + _AppService (not the full type name)
    [Inject] protected OfficeExport_AppService__Interface? Export_AppService { get; set; }
    [Inject] private DbProvider_AppService DbProvider_AppService { get; set; } = default!;   // member may match its type name — legal C# ("Color Color" rule)
    [Inject] private Layout_AppService Layout_AppService { get; set; } = default!;
    [Inject] protected WatcherEvent_AppService Watcher_AppService { get; set; } = default!;   // base "Watcher", not "WatcherEvent"

    // Third-party (Radzen) service members — TYPE keeps its name, MEMBER takes _AppService
    [Inject] protected DialogService Dialog_AppService { get; set; } = default!;
    [Inject] protected NotificationService Notification_AppService { get; set; } = default!;
}

// Local variables follow the same rule (camelCase base):
var dbProvider_AppService = new DbProvider_AppService { ProviderName = "InMemoryDb" };
```

Members typed as a genuinely non-service dependency (`NavigationManager`, `IJSRuntime JS`) keep their conventional names — the `_AppService` rule applies to members that hold an injected **service**.

---

## Parameter and Argument Naming

### Method Parameters

```csharp
// Clear names indicating type and purpose
public async Task<User_Model> CreateUserAsync(
    string firstName,
    string lastName,
    string email,
    DateTime dateOfBirth)
{
    // ...
}

// Avoid single letters (except in constrained scopes like loops)
// DON'T: CreateUserAsync(string f, string l, string e, DateTime d)
```

### Event Handler Naming

```csharp
// Event handlers follow {Source}_{Event} pattern
private void Button_Click()
{
    // Handle click
}

private async Task Form_Submit()
{
    // Handle submit
}

private void Input_Changed(string value)
{
    // Handle change
}
```

### Callback Naming

```csharp
[Parameter] public EventCallback<User_Model> OnSave { get; set; }
[Parameter] public EventCallback<Guid> OnDelete { get; set; }
[Parameter] public EventCallback OnCancel { get; set; }

// Call: await OnSave.InvokeAsync(user);
```

Naming:
- Callback parameters: `On{Action}` (e.g., `OnSave`, `OnDelete`, `OnCancel`)
- Return type in angle brackets indicates callback data type

---

## Constants and Enums

### Enum Naming

```csharp
public enum UserRoleEnum
{
    Admin = 1,
    Manager = 2,
    User = 3,
    Guest = 4
}

public enum StatusEnum
{
    Pending = 0,
    Active = 1,
    Inactive = 2,
    Deleted = 3
}
```

Naming:
- Enum names: `{Name}Enum`
- Enum values: `PascalCase`
- Numeric values 0-based or meaningful integers

### Constants

```csharp
// Class-level constants
public class UserConstants
{
    public const int MaxNameLength = 100;
    public const int MinPasswordLength = 8;
    public const string DefaultRole = "User";
}

// Magic numbers in code become constants
public const int MaxRetryAttempts = 3;
public const int TimeoutSeconds = 30;
```

Naming:
- Constant names: `PascalCase` for class-level (C# convention)
- Or `UPPER_CASE` for truly constant values
- Descriptive names indicating unit (e.g., `TimeoutSeconds`)

---

## Abbreviations and Acronyms

### When to Use Abbreviations

| Abbreviation | Use? | Example |
|--------------|------|---------|
| GUID | Yes | `userId`, `correlationId` |
| ID | Yes | `UserId`, `ProductId` |
| URL | Yes | `imageUrl`, `baseUrl` |
| HTTP | No | Use `Http` instead of `HTTP` |
| API | No | Use `Api` instead of `API` |
| JSON | No | Use `Json` instead of `JSON` |
| XML | No | Use `Xml` instead of `XML` |
| PDF | No | Use `Pdf` instead of `PDF` |
| SQL | No | Use `Sql` instead of `SQL` |
| JS/JavaScript | Yes | `jsRuntime`, `JSInterop` |
| CSS | Yes | `cssClass`, `CSSFramework` |
| HTML | No | Use `Html` instead of `HTML` |
| UI | No | Use `UserInterface` or full expansion |
| UX | No | Use `UserExperience` or full expansion |

### Acronym Casing Rules

- Single letter: lowercase (a, b, c)
- 2-letter acronym: all caps (ID, UI → except when lowercase needed like `jsRuntime`)
- 3+ letter acronym: PascalCase (Json, Html, Http)

Example:
```csharp
private IJSRuntime jsRuntime;  // IJSRuntime is an interface, jsRuntime is variable
public string HtmlContent { get; set; }
public Guid Id { get; set; }
public string ApiKey { get; set; }
```

---

## Special Naming Patterns

### Underscore Usage

```csharp
// Private backing field
private string _name;
public string Name { get; set; }

// Static/base class indicators
public abstract class Base__RadzenComponent { }

// CL entity files
public class _Row { }
public class RowLog_Row { }

// Extension methods file
public static class Enum_Extensions { }

// Utilities or services folder
public class _Services { }
public class _Utilities { }
```

### Plural vs Singular

```csharp
// Collections: plural
public List<User_Model> Users { get; set; }
public IEnumerable<Product_Model> Products { get; set; }

// Single items: singular
public User_Model User { get; set; }
public Product_Model Product { get; set; }

// Properties for UI lists: plural
[Parameter] public List<User_Model> Users { get; set; }

// Manager collection: plural
public IEnumerable<UserManager> UserManagers { get; set; }
```

### Boolean Property Naming

```csharp
// Use Is/Has/Can/Should prefixes
public bool IsActive { get; set; }
public bool HasPermission { get; set; }
public bool CanEdit { get; set; }
public bool ShouldValidate { get; set; }

// Methods returning bool
public bool IsUserActive(User_Model user)
public bool HasRequiredRole(User_Model user)
public bool CanDeleteUser(Guid userId)
```

---

## Summary Quick Reference

| Context | Pattern | Example |
|---------|---------|---------|
| C# Class | PascalCase | `UserManager` |
| C# Interface | `{Name}__Repo__Interface` / `{Name}_AppService__Interface` | `UserRepository__Repo__Interface`, `NewWindow_AppService__Interface` |
| C# Property | PascalCase | `FirstName` |
| C# Variable | camelCase | `firstName` |
| C# Constant | PascalCase | `MaxRetries` |
| Enum | PascalCaseEnum | `UserRoleEnum` |
| Entity Model | EntityName_Model | `User_Model` |
| Repository | EntityRepository | `UserRepository` |
| Manager | DomainManager | `UserManager` |
| Service (BL) | `{Name}__Service` | `CountriesNowSpaceManager__Service` |
| App Service (AL) | `{Name}_AppService` | `Layout_AppService` |
| Extension Class | Type_Extensions | `String_Extensions` |
| Folder | PascalCase | `Components`, `Managers` |
| Namespace | PascalCase segments | `JBC.ExploreTheWorld.DL.Users` |
| Razor Component | PascalCase | `UserForm.razor` |
| CSS Stylesheet | kebab-case | `app.css` |
| JavaScript File | kebab-case | `interop.js` |
