# Project Templates Document

## Overview

ExploreTheWorld projects are created using .NET project templates with specific configurations. This document defines the project file patterns (.csproj), package references, and configuration options used across all projects.

---

## Project File Structure Pattern

All projects follow a consistent .csproj structure:

```xml
<Project Sdk="Sdk.Type">
  <!-- Project Properties -->
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
  </PropertyGroup>

  <!-- Package References -->
  <ItemGroup>
    <PackageReference Include="Package.Name" Version="10.0.5" />
  </ItemGroup>

  <!-- Project References -->
  <ItemGroup>
    <ProjectReference Include="..\ReferencedProject\Project.csproj" />
  </ItemGroup>
</Project>
```

---

## SDK Types

### 1. Class Library SDK (Default)

Used for non-UI logic projects:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
  </PropertyGroup>
</Project>
```

**Applied to:**
- CL (Common Layer)
- DL (Dependency Layer)
- DL.MsJSInterop[.RevealJs], DL.MsOfficeApi.MsOfficeJs.{Word|Excel|PowerPoint}_Impl, DL.MsSystem, DL.MsSystemNet
- BL (Business Logic)
- AL (Application Layer)

---

### 2. Razor SDK (Class Library with Razor Support)

Used for Blazor component libraries:

```xml
<Project Sdk="Microsoft.NET.Sdk.Razor">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
    <AddRazorSupportForMvc>true</AddRazorSupportForMvc>
  </PropertyGroup>

  <ItemGroup>
    <FrameworkReference Include="Microsoft.AspNetCore.App" />
  </ItemGroup>
</Project>
```

**Applied to:**
- AL.BlazorLib
- AL.BlazorLib._radzen
- AL.BlazorLib.Server._radzen

---

### 3. Web SDK (ASP.NET Core Application)

Used for web servers and hosted applications:

```xml
<Project Sdk="Microsoft.NET.Sdk.Web">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
  </PropertyGroup>

  <ItemGroup>
    <FrameworkReference Include="Microsoft.AspNetCore.App" />
  </ItemGroup>
</Project>
```

**Applied to:**
- AL.BlazorWebApp (Hybrid server)

---

### 4. BlazorWebAssembly SDK

Used for client-side WebAssembly applications:

```xml
<Project Sdk="Microsoft.NET.Sdk.BlazorWebAssembly">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
  </PropertyGroup>
</Project>
```

**Applied to:**
- AL.BlazorWebApp.Client

---

## Package References by Layer

### CL (Common Layer)

**Packages:** None

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <!-- No ItemGroup - Foundation layer has no dependencies -->
</Project>
```

---

### DL (Dependency Layer)

**Packages:**
- Microsoft.Extensions.Logging.Abstractions (10.0.5)
- System.Text.Json (10.0.5)

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
    <PackageReference Include="System.Text.Json" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
  </ItemGroup>
</Project>
```

**Rationale:**
- Logging for DI-based logging in repository classes
- System.Text.Json for object serialization/deserialization

---

### DL.MsJSInterop (JavaScript Interop)

**Packages:**
- Microsoft.JSInterop (10.0.5)
- Same as DL

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.JSInterop" Version="10.0.5" />
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
    <ProjectReference Include="..\DL\ExploreTheWorld.DL.csproj" />
  </ItemGroup>
</Project>
```

---

### DL.MsSystem (System Utilities)

**Packages:**
- Microsoft.Extensions.Logging.Abstractions (10.0.5)

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
    <ProjectReference Include="..\DL\ExploreTheWorld.DL.csproj" />
  </ItemGroup>
</Project>
```

---

### DL.MsSystemNet (Network Utilities)

**Packages:**
- System.Net.Http.Json (10.0.5)
- Microsoft.Extensions.Logging.Abstractions (10.0.5)

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="System.Net.Http.Json" Version="10.0.5" />
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
    <ProjectReference Include="..\DL\ExploreTheWorld.DL.csproj" />
  </ItemGroup>
</Project>
```

---

### BL (Business Logic Layer)

**Packages:**
- Microsoft.Extensions.Logging.Abstractions (10.0.5)

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
    <ProjectReference Include="..\DL\ExploreTheWorld.DL.csproj" />
  </ItemGroup>
</Project>
```

---

### AL (Application Layer)

**Packages:**
- Microsoft.Extensions.Logging.Abstractions (10.0.5)

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
    <ProjectReference Include="..\BL\ExploreTheWorld.BL.csproj" />
    <ProjectReference Include="..\DL\ExploreTheWorld.DL.csproj" />
  </ItemGroup>
</Project>
```

---

### AL.BlazorLib (Reusable Blazor Components)

**Packages:**
- Radzen.Blazor (3.17.0+)
- Microsoft.AspNetCore.Components (10.0.5)
- Microsoft.AspNetCore.Components.Web (10.0.5)
- Microsoft.Extensions.Localization (10.0.5)
- Microsoft.Extensions.Logging.Abstractions (10.0.5)

```xml
<Project Sdk="Microsoft.NET.Sdk.Razor">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <AddRazorSupportForMvc>true</AddRazorSupportForMvc>
  </PropertyGroup>

  <ItemGroup>
    <FrameworkReference Include="Microsoft.AspNetCore.App" />
  </ItemGroup>

  <ItemGroup>
    <PackageReference Include="Radzen.Blazor" Version="3.17.0" />
    <PackageReference Include="Microsoft.AspNetCore.Components" Version="10.0.5" />
    <PackageReference Include="Microsoft.AspNetCore.Components.Web" Version="10.0.5" />
    <PackageReference Include="Microsoft.Extensions.Localization" Version="10.0.5" />
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\AL\ExploreTheWorld.AL.csproj" />
    <ProjectReference Include="..\BL\ExploreTheWorld.BL.csproj" />
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
    <ProjectReference Include="..\DL\ExploreTheWorld.DL.csproj" />
    <ProjectReference Include="..\DL.MsJSInterop\ExploreTheWorld.DL.MsJSInterop.csproj" />
    <ProjectReference Include="..\DL.MsSystemNet\ExploreTheWorld.DL.MsSystemNet.csproj" />
    <ProjectReference Include="..\DL.MsOfficeApi.OpenXml_Impl\ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.csproj" />
  </ItemGroup>
</Project>
```

---

### AL.BlazorLib._radzen (Radzen-Focused Variant)

**Same as AL.BlazorLib**

```xml
<Project Sdk="Microsoft.NET.Sdk.Razor">
  <!-- Same as AL.BlazorLib -->
</Project>
```

---

### AL.BlazorLib.Server._radzen (Server-Side Variant)

**Same as AL.BlazorLib**

```xml
<Project Sdk="Microsoft.NET.Sdk.Razor">
  <!-- Same as AL.BlazorLib -->
</Project>
```

---

### AL.BlazorWebApp (Hybrid Web Server)

**Packages:**
- Microsoft.AspNetCore.Components.WebAssembly.Server (10.0.5)
- Microsoft.AspNetCore.HeaderPropagation (10.0.5)
- Same logging and component packages

```xml
<Project Sdk="Microsoft.NET.Sdk.Web">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.AspNetCore.Components.WebAssembly.Server" Version="10.0.5" />
    <PackageReference Include="Microsoft.AspNetCore.HeaderPropagation" Version="10.0.5" />
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\AL.BlazorLib\ExploreTheWorld.AL.BlazorLib.csproj" />
    <ProjectReference Include="..\AL.BlazorWebApp.Client\ExploreTheWorld.AL.BlazorWebApp.Client.csproj" />
    <ProjectReference Include="..\AL\ExploreTheWorld.AL.csproj" />
    <ProjectReference Include="..\BL\ExploreTheWorld.BL.csproj" />
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
    <ProjectReference Include="..\DL\ExploreTheWorld.DL.csproj" />
    <ProjectReference Include="..\DL.MsJSInterop\ExploreTheWorld.DL.MsJSInterop.csproj" />
    <ProjectReference Include="..\DL.MsSystemNet\ExploreTheWorld.DL.MsSystemNet.csproj" />
  </ItemGroup>
</Project>
```

---

### AL.BlazorWebApp.Client (WebAssembly Client)

**Packages:**
- Microsoft.AspNetCore.Components.WebAssembly (10.0.5)

```xml
<Project Sdk="Microsoft.NET.Sdk.BlazorWebAssembly">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.AspNetCore.Components.WebAssembly" Version="10.0.5" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\AL.BlazorLib\ExploreTheWorld.AL.BlazorLib.csproj" />
    <ProjectReference Include="..\AL\ExploreTheWorld.AL.csproj" />
    <ProjectReference Include="..\BL\ExploreTheWorld.BL.csproj" />
    <ProjectReference Include="..\CL\ExploreTheWorld.CL.csproj" />
    <ProjectReference Include="..\DL\ExploreTheWorld.DL.csproj" />
    <ProjectReference Include="..\DL.MsJSInterop\ExploreTheWorld.DL.MsJSInterop.csproj" />
    <ProjectReference Include="..\DL.MsSystemNet\ExploreTheWorld.DL.MsSystemNet.csproj" />
  </ItemGroup>
</Project>
```

---

### DL.{Name}Api (External REST API Client)

**Purpose:** Self-contained REST API integration — no BL dependency.
**IMPORTANT:** Do NOT add `<ImplicitUsings>enable</ImplicitUsings>` — required for ._netF shared-link compatibility.

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

**Folder structure:**
```
DL.{Name}Api/
  _Interfaces/{Name}Api_Interface.cs
  _Rows/{Entity}_Row.cs
  {Name}Api__Repo.cs
```

---

### DL.{Name}Api._netF (.NET Framework Shared-Link)

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net481</TargetFramework>
    <LangVersion>latest</LangVersion>
    <Nullable>enable</Nullable>
    <AssemblyName>JBC.ExploreTheWorld.DL.{Name}Api._netF</AssemblyName>
    <RootNamespace>JBC.ExploreTheWorld.DL.{Name}Api</RootNamespace>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="..\CL._netF\ExploreTheWorld.CL._netF.csproj" />
    <ProjectReference Include="..\DL._netF\ExploreTheWorld.DL._netF.csproj" />
  </ItemGroup>

  <ItemGroup>
    <PackageReference Include="System.Net.Http" Version="4.3.4" />
    <PackageReference Include="System.Net.Http.Json" Version="8.0.1" />
    <PackageReference Include="System.Text.Json" Version="8.0.5" />
  </ItemGroup>

  <ItemGroup>
    <Compile Include="..\DL.{Name}Api\**\*.cs">
      <Link>%(RecursiveDir)%(Filename)%(Extension)</Link>
    </Compile>
  </ItemGroup>

  <ItemGroup>
    <Compile Remove="..\DL.{Name}Api\obj\**" />
    <Compile Remove="obj\**" />
    <EmbeddedResource Remove="obj\**" />
    <None Remove="obj\**" />
  </ItemGroup>
</Project>
```

---

### AL.WinFormApp (BlazorWebView WinForms Host — net10.0)

```xml
<Project Sdk="Microsoft.NET.Sdk.Razor">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net10.0-windows</TargetFramework>
    <Nullable>enable</Nullable>
    <UseWindowsForms>true</UseWindowsForms>
    <ImplicitUsings>enable</ImplicitUsings>
    <AssemblyName>JBC.ExploreTheWorld.AL.WinFormApp</AssemblyName>
    <RootNamespace>JBC.ExploreTheWorld.AL.WinFormApp</RootNamespace>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.AspNetCore.Components.WebView.WindowsForms" Version="10.0.51" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\AL.BlazorLib\ExploreTheWorld.AL.BlazorLib.csproj" />
    <!-- Add other project refs as needed -->
  </ItemGroup>
</Project>
```

**Program.cs pattern:**
```csharp
[STAThread]
static void Main()
{
    var services = new ServiceCollection();
    services.AddWindowsFormsBlazorWebView();
#if DEBUG
    services.AddBlazorWebViewDeveloperTools();
#endif
    services.AddRadzenComponents();
    services.AddTransient<SomeApi_Interface, SomeApi__Repo>();
    var serviceProvider = services.BuildServiceProvider();
    ApplicationConfiguration.Initialize();
    var restService = serviceProvider.GetRequiredService<SomeRestApi_Interface>();
    var countriesNowService = serviceProvider.GetRequiredService<SomeCountriesNowApi_Interface>();
    Application.Run(new Main_Form(restService, countriesNowService));
}
```

---

### AL.WinFormApp._netF (Traditional WinForms — net481)

No BlazorWebView. Uses DataGridView controls and direct service instantiation.

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net481</TargetFramework>
    <LangVersion>latest</LangVersion>
    <Nullable>enable</Nullable>
    <UseWindowsForms>true</UseWindowsForms>
    <AssemblyName>JBC.ExploreTheWorld.AL.WinFormApp._netF</AssemblyName>
    <RootNamespace>JBC.ExploreTheWorld.AL.WinFormApp</RootNamespace>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="..\CL._netF\ExploreTheWorld.CL._netF.csproj" />
    <ProjectReference Include="..\DL._netF\ExploreTheWorld.DL._netF.csproj" />
    <!-- Add API ._netF project refs as needed -->
  </ItemGroup>

  <!-- LINK child form source files from AL.WinFormApp (shared-link compilation pattern) -->
  <ItemGroup>
    <Compile Include="..\AL.WinFormApp\_Forms\RestCountriesCom_Form.cs" Link="_Forms\RestCountriesCom_Form.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\RestCountriesCom_Form.Designer.cs" Link="_Forms\RestCountriesCom_Form.Designer.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\CountriesNowSpace_Form.cs" Link="_Forms\CountriesNowSpace_Form.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\CountriesNowSpace_Form.Designer.cs" Link="_Forms\CountriesNowSpace_Form.Designer.cs" />
  </ItemGroup>
</Project>
```

**Program.cs pattern:**
```csharp
[STAThread]
static void Main()
{
    Application.EnableVisualStyles();
    Application.SetCompatibleTextRenderingDefault(false);
    var apiService = new SomeApi__Repo();
    Application.Run(new Main_Form(apiService));
}
```

---

## .NET Framework Variant Projects

All ._netF variants use the **shared-link compilation pattern**. They link source files from their net10.0 counterparts without duplication.

### Example: CL._netF

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net481</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <!-- LINK source files from CL project (no duplication) -->
  <ItemGroup>
    <Compile Include="..\CL\**\*.cs" />
  </ItemGroup>

  <!-- EXCLUDE build artifacts -->
  <ItemGroup>
    <Compile Remove="obj/**" />
    <Compile Remove="bin/**" />
  </ItemGroup>
</Project>
```

**._netF Compatibility Notes:**
- Packages: System.Net.Http (4.3.4) for Framework compatibility
- Logging: Microsoft.Extensions.Logging.Abstractions (if used)
- No Blazor or Web-specific packages
- Applied to: CL._netF, DL._netF, BL._netF, AL._netF

---

## Common Configuration Settings

### PropertyGroup Settings

```xml
<PropertyGroup>
  <!-- Framework Target -->
  <TargetFramework>net10.0</TargetFramework>

  <!-- Enable implicit using statements (using System, etc.) -->
  <ImplicitUsings>enable</ImplicitUsings>

  <!-- Enable nullable reference types -->
  <Nullable>enable</Nullable>

  <!-- Use latest C# version -->
  <LangVersion>latest</LangVersion>

  <!-- Optional: Assembly version -->
  <AssemblyVersion>1.0.0.0</AssemblyVersion>
  <FileVersion>1.0.0.0</FileVersion>

  <!-- Optional: Generate documentation -->
  <GenerateDocumentationFile>true</GenerateDocumentationFile>

  <!-- Optional: Package info for NuGet -->
  <Authors>Your Name</Authors>
  <Description>Package description</Description>
  <Version>1.0.0</Version>
</PropertyGroup>
```

### Build Optimization

For Release builds, add:

```xml
<PropertyGroup Condition="'$(Configuration)|$(Platform)'=='Release|AnyCPU'">
  <DebugType>none</DebugType>
  <DebugSymbols>false</DebugSymbols>
  <DebugInfo>None</DebugInfo>
</PropertyGroup>
```

---

## Package Version Strategy

### Versioning

All packages use **major.minor.patch** semantic versioning:
- **10.0.5** - .NET 10.0 release packages
- **4.3.4** - .NET Framework compatibility packages
- **3.17.0+** - Radzen Blazor (matches Radzen versions)

### Package Updates

When updating packages:
1. Update all related packages to compatible versions
2. Test build against all solution files
3. Verify no breaking changes in APIs used

Example compatible sets:
- ASP.NET Core 10.0: All Microsoft.AspNetCore.* at 10.0.x
- Extensions 10.0: All Microsoft.Extensions.* at 10.0.x
- System packages: All System.* at 10.0.x

---

## Creating New Projects

### Template Steps

1. **Determine project type** - Use SDK table above
2. **Create folder** - `mkdir src\NewProjectName`
3. **Create .csproj** - Use template from this document
4. **Update project references** - Follow dependency rules
5. **Add to solution files** - Include in appropriate .sln
6. **Create folder structure** - Follow FILE_STRUCTURE_GUIDE.md
7. **Create _Imports.razor** - For Razor projects
8. **Test build** - Run `dotnet build src\SolutionName.sln`

### Minimal .csproj Template

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
  </PropertyGroup>

  <ItemGroup>
    <!-- Add package references as needed -->
  </ItemGroup>

  <ItemGroup>
    <!-- Add project references as needed -->
  </ItemGroup>
</Project>
```

---

## Office Web Add-in Projects (Blazor WASM)

Each Office host (Word, Excel, PowerPoint) has **two sibling projects** under `code-JBC-ExploreTheWorld\`:

### Manifest Project (net481 old-style csproj)

**Project type:** `{C1CDDADD-2546-481F-9697-4EA41081F2FC};{14822709-B5A1-4724-98CA-57A101D1B079};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}`  
**Imports:** `Microsoft.VisualStudio.SharePoint.targets`  
**Key element:** `<CoreWebProject>` pointing to the sibling Blazor WASM csproj  
**Manifest XML** lives in `{ProjectName}Manifest\` subfolder  
**Copy manifest** to `OfficeAddins/{Host}/Web/` for sideloading

### Blazor WASM Project (net10.0)

**SDK:** `Microsoft.NET.Sdk.BlazorWebAssembly`  
**AssemblyName / RootNamespace:** `ExploreTheWorld.AL.{ProjectName}`  
**Packages (v10.0.5):** `Microsoft.AspNetCore.Components.WebAssembly`, `.DevServer`, `.JSInterop.WebAssembly`, `System.Text.Json`

**wwwroot key files:**
- `index.html` — loads `office.js` from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`
- `{AssemblyName}.lib.module.js` — Blazor WASM lifecycle hooks (`beforeStart`/`afterStarted`), calls `Office.onReady()`
- `Commands/commands.js` — toggle taskpane via `Office.addin.showAsTaskpane()`/`Office.addin.hide()` + `Office.actions.associate("toggleTaskpane", ...)`
- `css/app.css` — task pane styles
- `Images/Button{16,32,64,80}x{16,32,64,80}.png` — Ribbon button icons

**Pages:**
- `Index.razor` / `.cs` / `.js` — Document/Workbook/Presentation info via `Word.run()` / `Excel.run()` / `PowerPoint.run()`
- `Events.razor` / `.cs` / `.js` — Event watcher using `[JSInvokable] OnEventLogged(string eventName, string timestamp)` callback from JS

**JS interop pattern:** ES module import in `OnAfterRenderAsync(firstRender)`:
```csharp
_jsModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/PageName.razor.js");
```

**SharedRuntime requirement:** Manifest must use `<Runtime resid="...Shared.Url" lifetime="long">` so `ExecuteFunction` (commands.js) and task pane share the same runtime.

---

## Test Projects

### net10.0 Test Projects

All net10.0 test projects use `xunit 2.9.x`, `FluentAssertions 8.x`, and `Microsoft.NET.Test.Sdk 18.x`.

| Project | TFM | Extra packages |
|---------|-----|----------------|
| `UnitTests` | `net10.0` | `Moq 4.x`, `coverlet.collector` |
| `IntegrationTests` | `net10.0` | `Microsoft.EntityFrameworkCore.InMemory`, `Microsoft.AspNetCore.Mvc.Testing` |
| `OpenXmlLibTests` | `net10.0` | `coverlet.collector` (OpenXml SDK flows from the `DL.MsOfficeApi.OpenXml_Impl` project reference) |
| `WinFormAppTests` | `net10.0-windows` | `FlaUI.Core 5.x`, `FlaUI.UIA3 5.x` |
| `OfficeAddinTests` | `net10.0-windows` | `FlaUI.Core 5.x`, `FlaUI.UIA3 5.x`, NetOffice project refs |
| `AccessDbTests` | `net10.0-windows` | `FlaUI.Core 5.x`, `FlaUI.UIA3 5.x` (drives MSACCESS.EXE) |

### .NET Framework Test Projects (`._netF`)

All `._netF` test projects use `xunit 2.9.x`, **`FluentAssertions 6.12.0`** (last version with .NET Framework support), and `Microsoft.NET.Test.Sdk 17.9.0`.

```xml
<!-- UnitTests._netF — links all source from UnitTests -->
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net481</TargetFramework>
    <LangVersion>latest</LangVersion>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <AssemblyName>JBC.ExploreTheWorld.UnitTests._netF</AssemblyName>
    <RootNamespace>JBC.ExploreTheWorld.UnitTests</RootNamespace>
    <IsTestProject>true</IsTestProject>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.9.0" />
    <PackageReference Include="xunit" Version="2.9.3" />
    <PackageReference Include="xunit.runner.visualstudio" Version="3.1.5">
      <IncludeAssets>runtime; build; native; contentfiles; analyzers; buildtransitive</IncludeAssets>
      <PrivateAssets>all</PrivateAssets>
    </PackageReference>
    <PackageReference Include="Moq" Version="4.20.72" />
    <PackageReference Include="FluentAssertions" Version="6.12.0" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\CL._netF\ExploreTheWorld.CL._netF.csproj" />
    <!-- ... other _netF layer refs ... -->
  </ItemGroup>
  <!-- Link shared test source from UnitTests (shared-link compilation pattern). -->
  <ItemGroup>
    <Compile Include="..\UnitTests\**\*.cs">
      <Link>%(RecursiveDir)%(Filename)%(Extension)</Link>
    </Compile>
  </ItemGroup>
  <ItemGroup>
    <Compile Remove="..\UnitTests\obj\**" />
    <Compile Remove="..\UnitTests\bin\**" />
    <Compile Remove="obj\**" />
  </ItemGroup>
</Project>
```

**IntegrationTests._netF** is **standalone** (no shared-link): `IDbContextFactory<T>` (EF Core 5+) and `ExecuteDeleteAsync()` (EF Core 7+) are unavailable in EF Core 3.x. The _netF repo uses `Func<ExploreTheWorldDbContext>` instead. Add `Microsoft.EntityFrameworkCore.InMemory Version="3.1.32"`.

**WinFormAppTests._netF** is **standalone**: the exe path and button names differ from the net10.0 WinForms app. Set `ETW_WINFORMAPP_NETF_PATH` env var to override the exe location.

**OfficeAddinTests._netF** links all source from `OfficeAddinTests` and replaces the NetOffice project references with **NuGet packages**:

```xml
<PackageReference Include="NetOfficeFw.Core"       Version="1.9.9" />
<PackageReference Include="NetOfficeFw.Office"     Version="1.9.9" />
<PackageReference Include="NetOfficeFw.Word"       Version="1.9.9" />
<PackageReference Include="NetOfficeFw.Excel"      Version="1.9.9" />
<PackageReference Include="NetOfficeFw.PowerPoint" Version="1.9.9" />
```

---

## Troubleshooting Project Issues

### Issue: Build fails with "project not found"
**Solution:** Verify ProjectReference paths are relative and correct

### Issue: NuGet packages conflict
**Solution:** Ensure all package versions are compatible across layers

### Issue: Circular dependency
**Solution:** Review PROJECT_TEMPLATES.md layer dependency rules

### Issue: "_Imports.razor not found" in Razor project
**Solution:** Create _Imports.razor in project root with global usings

### Issue: Type not accessible across projects
**Solution:** Add ProjectReference and ensure namespace is correct
