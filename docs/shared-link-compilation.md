# Shared-Link Compilation Guide

## Overview

The ExploreTheWorld architecture uses **shared-link compilation** to support both .NET 10.0 (net10.0) and .NET Framework 4.8.1 (net481) targets **without duplicating code**. This approach allows 100% code reuse across framework targets while maintaining separate project files and NuGet package configurations.

---

## What is Shared-Link Compilation?

Shared-link compilation is a .NET feature that allows one project to **link source files** from another project's folder structure without physically copying them. The source files are compiled into the linking project as if they were directly in that project.

### Benefits

1. **Zero Code Duplication** - Single source of truth for all business logic
2. **Simplified Maintenance** - Bug fixes apply to all frameworks automatically
3. **Consistent Behavior** - Business logic identical across net10.0 and net481
4. **Framework Flexibility** - Different dependencies per framework version as needed
5. **Clean Build Output** - Each framework produces its own compiled assembly

### When to Use

- **Shared-link projects (._netF)** - Link source from framework-agnostic projects
- **Not for UI projects** - Blazor components framework-specific, no ._netF variant
- **Not for web projects** - ASP.NET Core is .NET Core only

---

## Implementation Pattern

### Net10.0 Project (Source)

```
CL/
├── ExploreTheWorld.CL.csproj          # net10.0 target
├── Enum/
│   └── Enum_Extensions.cs
├── TreeStructure/
│   └── TreeNode.cs
└── _Row/
    ├── _Row.cs
    ├── RowLog_Row.cs
    └── ColumnLog_Row.cs
```

```xml
<!-- ExploreTheWorld.CL.csproj -->
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
</Project>
```

### .NET Framework Variant (Linking Project)

```
CL._netF/
├── ExploreTheWorld.CL._netF.csproj    # net481 target
└── (no source files - all linked from CL/)
```

```xml
<!-- ExploreTheWorld.CL._netF.csproj -->
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net481</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <!-- LINK all source files from net10.0 project -->
  <ItemGroup>
    <Compile Include="..\CL\**\*.cs" />
  </ItemGroup>

  <!-- EXCLUDE build artifacts from link -->
  <ItemGroup>
    <Compile Remove="obj/**" />
    <Compile Remove="bin/**" />
  </ItemGroup>
</Project>
```

---

## Glob Pattern Explanation

### Pattern: `..\CL\**\*.cs`

```
..\CL\          - Parent directory, then CL folder (relative path)
  **\           - Any folder at any depth (recursive)
    *.cs        - All C# source files
```

Examples matched:
- `..\CL\Enum_Extensions.cs` ✓
- `..\CL\Enum\Enum_Extensions.cs` ✓
- `..\CL\_Row\_Row.cs` ✓
- `..\CL\nested\deep\folder\AnyFile.cs` ✓

---

## Excluding Build Artifacts

Build artifacts must be excluded to prevent compilation errors:

```xml
<ItemGroup>
  <Compile Include="..\CL\**\*.cs" />
</ItemGroup>

<ItemGroup>
  <!-- Remove obj/ and bin/ folders from the link -->
  <Compile Remove="obj/**" />
  <Compile Remove="bin/**" />
</ItemGroup>
```

**Why exclude?**
- `obj/` folder: Contains intermediate build files
- `bin/` folder: Contains compiled assemblies
- These are auto-generated and shouldn't be recompiled

---

## Projects Using Shared-Link Compilation

### Test Projects

| `._netF` Test Project | Source Linked From | Why standalone instead of linked |
|----------------------|--------------------|----------------------------------|
| `UnitTests._netF` | All of `UnitTests/` | — (fully linked) |
| `IntegrationTests._netF` | None — standalone | `IDbContextFactory<T>` (EF Core 5+) and `ExecuteDeleteAsync()` (EF Core 7+) unavailable in EF Core 3.x; fixture and tests rewritten for `Func<DbContext>` factory |
| `SqliteDbTests._netF` | None — standalone | No net10.0 counterpart; targets the EF Core 3.x Sqlite provider and the `_netF` repository (`CreateDbContext()` only) |
| `OpenXmlLibTests._netF` | All of `OpenXmlLibTests/` | — (fully linked); OpenXml SDK API is framework-agnostic |
| `WinFormAppTests._netF` | None — standalone | Different exe path (`net481/` vs `net10.0-windows/`) and different button labels in `Main_Form` |
| `OfficeAddinTests._netF` | All of `OfficeAddinTests/` | — (fully linked); FlaUI API is framework-agnostic |

**FluentAssertions version:** `6.12.0` for all `._netF` test projects (FluentAssertions 7+ requires .NET 6+).

---

### 1. CL._netF (Common Layer - Framework Support)

**Source:** `CL/`
**Target Framework:** net481

```xml
<ItemGroup>
  <Compile Include="..\CL\**\*.cs" />
</ItemGroup>
<ItemGroup>
  <Compile Remove="obj/**" />
  <Compile Remove="bin/**" />
</ItemGroup>
```

**Use Case:** All CL utilities (enums, tree nodes, entities, logging models) available to .NET Framework applications

---

### 2. DL._netF (Dependency Layer - Framework Support)

**Source:** `DL/`
**Target Framework:** net481
**Additional Package:** System.Net.Http (4.3.4) for Framework compatibility

```xml
<ItemGroup>
  <Compile Include="..\DL\**\*.cs" />
</ItemGroup>
<ItemGroup>
  <Compile Remove="obj/**" />
  <Compile Remove="bin/**" />
</ItemGroup>

<!-- Add Framework-specific compatibility package -->
<ItemGroup>
  <PackageReference Include="System.Net.Http" Version="4.3.4" />
  <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
</ItemGroup>

<ItemGroup>
  <ProjectReference Include="..\CL._netF\ExploreTheWorld.CL._netF.csproj" />
</ItemGroup>
```

---

### 3. BL._netF (Business Logic - Framework Support)

**Source:** `BL/`
**Target Framework:** net481

```xml
<ItemGroup>
  <Compile Include="..\BL\**\*.cs" />
</ItemGroup>
<ItemGroup>
  <Compile Remove="obj/**" />
  <Compile Remove="bin/**" />
</ItemGroup>

<ItemGroup>
  <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
</ItemGroup>

<ItemGroup>
  <ProjectReference Include="..\CL._netF\ExploreTheWorld.CL._netF.csproj" />
  <ProjectReference Include="..\DL._netF\ExploreTheWorld.DL._netF.csproj" />
</ItemGroup>
```

---

### 4. AL._netF (Application Layer - Framework Support)

**Source:** `AL/`
**Target Framework:** net481

```xml
<ItemGroup>
  <Compile Include="..\AL\**\*.cs" />
</ItemGroup>
<ItemGroup>
  <Compile Remove="obj/**" />
  <Compile Remove="bin/**" />
</ItemGroup>

<ItemGroup>
  <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.5" />
</ItemGroup>

<ItemGroup>
  <ProjectReference Include="..\CL._netF\ExploreTheWorld.CL._netF.csproj" />
  <ProjectReference Include="..\BL._netF\ExploreTheWorld.BL._netF.csproj" />
  <ProjectReference Include="..\DL._netF\ExploreTheWorld.DL._netF.csproj" />
</ItemGroup>
```

---

### 5. AL.WinFormApp._netF (WinForms Traditional App — Partial Shared-Link)

**Source:** `AL.WinFormApp/_Forms/` (selected files)
**Target Framework:** net481

`AL.WinFormApp._netF` partially uses shared-link compilation. Child forms and all Watcher/Export source files are linked from `AL.WinFormApp`. `Main_Form` is **not** linked — `._netF` has its own 3-button `Main_Form`. `WebView_Form.*` is **not** linked because it depends on `BlazorWebView` and `Microsoft.AspNetCore.Components.WebView.WindowsForms`, which are .NET 10 only.

```xml
  <!-- LINK child form and Watcher/Export source files from AL.WinFormApp (shared-link compilation pattern) -->
  <ItemGroup>
    <Compile Include="..\AL.WinFormApp\_Forms\RestCountriesCom_Form.cs" Link="_Forms\RestCountriesCom_Form.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\RestCountriesCom_Form.Designer.cs" Link="_Forms\RestCountriesCom_Form.Designer.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\CountriesNowSpace_Form.cs" Link="_Forms\CountriesNowSpace_Form.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\CountriesNowSpace_Form.Designer.cs" Link="_Forms\CountriesNowSpace_Form.Designer.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Export\ExportType_Enum.cs" Link="_Forms\_Export\ExportType_Enum.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Export\ExportMethod_Enum.cs" Link="_Forms\_Export\ExportMethod_Enum.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Export\ExportMenuHelper.cs" Link="_Forms\_Export\ExportMenuHelper.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Export\ExportLog_Form.cs" Link="_Forms\_Export\ExportLog_Form.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Export\ExportLog_Form.Designer.cs" Link="_Forms\_Export\ExportLog_Form.Designer.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\WatcherInteropMethod_Enum.cs" Link="_Forms\_Watcher\WatcherInteropMethod_Enum.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsOfficeEvent_Record.cs" Link="_Forms\_Watcher\MsOfficeEvent_Record.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsOfficeEvents_Repo.cs" Link="_Forms\_Watcher\MsOfficeEvents_Repo.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsOfficeJsonWriter_Helper.cs" Link="_Forms\_Watcher\MsOfficeJsonWriter_Helper.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\WatcherComHelper.cs" Link="_Forms\_Watcher\WatcherComHelper.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsWord_Watcher_Form.cs" Link="_Forms\_Watcher\MsWord_Watcher_Form.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsWord_Watcher_Form.Designer.cs" Link="_Forms\_Watcher\MsWord_Watcher_Form.Designer.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsExcel_Watcher_Form.cs" Link="_Forms\_Watcher\MsExcel_Watcher_Form.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsExcel_Watcher_Form.Designer.cs" Link="_Forms\_Watcher\MsExcel_Watcher_Form.Designer.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsPowerPoint_Watcher_Form.cs" Link="_Forms\_Watcher\MsPowerPoint_Watcher_Form.cs" />
    <Compile Include="..\AL.WinFormApp\_Forms\_Watcher\MsPowerPoint_Watcher_Form.Designer.cs" Link="_Forms\_Watcher\MsPowerPoint_Watcher_Form.Designer.cs" />
  </ItemGroup>
```

`.resx` files for each form are kept in their respective `AL.WinFormApp._netF/_Forms/` subfolder (not linked) because WinForms `EmbeddedResource`/`DependentUpon` linking adds complexity for identical default boilerplate.

**NetOffice in the Watcher forms:**

The shared Watcher form source (`MsWord/Excel/PowerPoint_Watcher_Form.cs`) uses NetOffice types unconditionally — no `#if NETFRAMEWORK` guards. NetOffice is available on both targets through different supply paths:

- **net481 (`._netF`):** `NetOfficeFw.Core/Excel/Word/PowerPoint` NuGet packages (v1.9.9)
- **net10.0 (`AL.WinFormApp`):** `NetOffice__10` project references from `code-zgh-NetOfficeFw__NetOffice__10`

`WatcherComHelper.GetActiveCom(progId)` handles the platform difference for `GetActiveObject` internally (its `#if NETFRAMEWORK` guard is an implementation detail of the helper, not the consumer forms).

**Resulting `_Forms/` layout:**

| File | Location | Notes |
|------|----------|-------|
| `WebView_Form.*` | `AL.WinFormApp/_Forms/` | net10.0 only — BlazorWebView host |
| `Main_Form.*` | `AL.WinFormApp/_Forms/` | net10.0 version (no Watcher button) |
| `Main_Form.*` | `AL.WinFormApp._netF/_Forms/` | net481 version — NOT linked; has Watcher button |
| `RestCountriesCom_Form.cs/.Designer.cs` | `AL.WinFormApp/_Forms/` *(canonical)* | Linked into `._netF` |
| `CountriesNowSpace_Form.cs/.Designer.cs` | `AL.WinFormApp/_Forms/` *(canonical)* | Linked into `._netF` |
| `_Export/ExportType_Enum.cs` etc. | `AL.WinFormApp/_Forms/_Export/` *(canonical)* | Linked into `._netF` |
| `_Watcher/MsWord_Watcher_Form.cs` etc. | `AL.WinFormApp/_Forms/_Watcher/` *(canonical)* | Linked into `._netF` |
| `*.resx` | `AL.WinFormApp._netF/_Forms/` | Not linked — kept for simplicity |

---

## Project Dependency Graph

### Net10.0 Projects (Primary)

```
AL.BlazorWebApp     (Web SDK, net10.0)
     ↓
AL.BlazorLib        (Razor SDK, net10.0)
     ↓
AL                  (Class Library, net10.0)
     ↓
BL                  (Class Library, net10.0)
     ↓
DL                  (Class Library, net10.0)
     ↓
CL                  (Class Library, net10.0)
```

### Net481 Projects (Shared-Link)

```
AL._netF            (Class Library, net481) ← Links from AL/
     ↓                    ↕
BL._netF            (Class Library, net481) ← Links from BL/
     ↓                    ↕
DL._netF            (Class Library, net481) ← Links from DL/
     ↓                    ↕
CL._netF            (Class Library, net481) ← Links from CL/
```

**Note:** ._netF projects have **no UI variants** (no server or client blazor)

---

## Building with Shared-Link Projects

### Individual Project Build

```bash
# Build just the net481 variant
dotnet build src\DL._netF\ExploreTheWorld.DL._netF.csproj
```

### Solution Build

```bash
# Build entire solution including ._netF projects
dotnet build src\JBC.ExploreTheWorld.sln

# Build only ._netF solution
dotnet build src\JBC.ExploreTheWorld._netF.sln
```

### Build Process

When building a ._netF project:

1. **Resolve source files** - Find all `*.cs` files in source project folder
2. **Link files** - Import them as if they were in the ._netF project
3. **Remove exclusions** - Skip `obj/` and `bin/` folders
4. **Compile** - Use net481-compatible compiler settings
5. **Output assembly** - Generate `*.dll` for .NET Framework 4.8.1

---

## Common Issues and Solutions

### Issue 1: "Multiple Files Found" Error

**Symptom:** Compiler error about duplicate type definitions

**Cause:** Accidental duplication of source files in both folders

**Solution:**
```xml
<!-- Ensure no source files in ._netF project folder -->
<!-- ._netF folder should contain ONLY the .csproj file -->
```

### Issue 2: Build Artifacts Included in Link

**Symptom:** Compiler errors about obj/ or bin/ files

**Solution:**
```xml
<ItemGroup>
  <Compile Remove="obj/**" />
  <Compile Remove="bin/**" />
</ItemGroup>
```

### Issue 3: Path Not Found

**Symptom:** "Cannot find path '..\SourceProject\...'"

**Cause:** Incorrect relative path in Compile Include

**Solution:** Verify folder structure matches relative path:
```
code-JBC-ExploreTheWorld/
├── CL/                    (Source)
└── CL._netF/              (Linking project)
    └── In CL._netF.csproj: Include="..\CL\**\*.cs" ✓
```

### Issue 4: IntelliSense Not Working

**Symptom:** IDE doesn't show source files from linked project

**Solution:**
- Visual Studio: Clean and rebuild solution
- VS Code: Reload workspace
- Usually resolves after first successful build

---

## Framework-Specific Code

### When Source Must Differ Between Frameworks

Use conditional compilation if needed:

```csharp
// In shared source file (CL/Enum_Extensions.cs)

#if NET10_0_OR_GREATER
    // .NET 10.0 specific code
    private static readonly JsonSerializerOptions _jsonOptions = new() { };
#else
    // .NET Framework 4.8.1 code
    private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions();
#endif
```

**Best Practice:** Minimize framework-specific code. Use abstractions instead.

---

## Adding New Shared-Link Projects

### Step-by-Step

1. **Create source project** (net10.0)
   ```bash
   mkdir NewProject
   # Create NewProject/ExploreTheWorld.NewProject.csproj
   ```

2. **Add content** to source project
   ```
   NewProject/
   ├── ExploreTheWorld.NewProject.csproj
   ├── MyClass.cs
   └── Subfolder/
       └── AnotherClass.cs
   ```

3. **Create ._netF variant**
   ```bash
   mkdir NewProject._netF
   # Create NewProject._netF/ExploreTheWorld.NewProject._netF.csproj
   ```

4. **Configure .csproj**
   ```xml
   <Project Sdk="Microsoft.NET.Sdk">
     <PropertyGroup>
       <TargetFramework>net481</TargetFramework>
     </PropertyGroup>

     <ItemGroup>
       <Compile Include="..\NewProject\**\*.cs" />
     </ItemGroup>
     <ItemGroup>
       <Compile Remove="obj/**" />
       <Compile Remove="bin/**" />
     </ItemGroup>

     <ItemGroup>
       <!-- Add ._netF variant of dependencies -->
       <ProjectReference Include="..\DependencyLayer._netF\..." />
     </ItemGroup>
   </Project>
   ```

5. **Add to solution file**
   ```
   src\JBC.ExploreTheWorld._netF.sln
   ```

6. **Test build**
   ```bash
   dotnet build src\JBC.ExploreTheWorld._netF.sln
   ```

---

## Deployment Scenarios

### .NET 10.0 Only

Deploy main solution projects:
- Use `JBC.ExploreTheWorld.sln`
- Deploy compiled `DL.dll`, `BL.dll`, `AL.dll`, etc.
- Include `AL.BlazorWebApp.dll` for web hosting

### .NET Framework Support

Deploy framework variants:
- Use `JBC.ExploreTheWorld._netF.sln`
- Deploy compiled `DL._netF.dll`, `BL._netF.dll`, `AL._netF.dll`
- No Blazor/Web components available in Framework target

### Both Frameworks (Side-by-Side)

Deploy both sets:
```
Distribution/
├── net10.0/
│   ├── DL.dll
│   ├── BL.dll
│   └── AL.dll
└── net481/
    ├── DL._netF.dll
    ├── BL._netF.dll
    └── AL._netF.dll
```

Applications can choose which version to reference at build time.

---

## Performance Considerations

### Build Time

- **First build:** Same as standard projects (source files compiled)
- **Subsequent builds:** Same as standard projects (incremental compilation)
- **Shared-link overhead:** Negligible (file linking is fast)

### Runtime Performance

- **No runtime overhead** - All code compiled to native .NET IL
- **Assembly size:** Identical to duplicated source approach
- **Memory usage:** Identical to duplicated source approach

### Optimization

Shared-link compilation has **zero performance cost** compared to code duplication.

---

## Best Practices

### 1. Keep Separation Clean

✓ Do: Only source files in source project, only project file in ._netF
✗ Don't: Keep any source files in ._netF folder

### 2. Match Project Structure

✓ Do: Mirror folder structure exactly
✗ Don't: Create different folder structures in linking project

### 3. Use Proper Glob Patterns

✓ Do: `Include="..\SourceProject\**\*.cs"`
✗ Don't: List individual files

### 4. Exclude Build Artifacts

✓ Do: Always use `Remove="obj/**"` and `Remove="bin/**"`
✗ Don't: Forget to exclude build folders

### 5. Reference ._netF Dependencies

✓ Do: In DL._netF, reference CL._netF (not CL)
✗ Don't: Mix framework versions in project references

### 6. Test Both Frameworks

```bash
# Build both framework variants
dotnet build src\JBC.ExploreTheWorld.sln         # net10.0
dotnet build src\JBC.ExploreTheWorld._netF.sln   # net481
```

---

## Documentation

Shared-link compilation requirements are documented in:
- **ARCHITECTURE_STANDARDS.md** - Framework support strategy
- **PROJECT_TEMPLATES.md** - Project configuration templates
- This guide - Implementation details

Reference these when:
- Creating new ._netF projects
- Troubleshooting build issues
- Deciding on framework targets
