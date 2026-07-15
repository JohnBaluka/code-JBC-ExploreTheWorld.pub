# TestResults

Visual artifacts produced by the **UI test projects** (`src/Tests/*`). Each UI test drives a
real running app — a desktop window via **FlaUI** (UIA3) or a browser page via **Playwright**
(Chromium) — and captures screenshots (and, for browser tests, a video) so a reviewer can *see*
what the test saw without re-running it.

These files are **committed to the repository** as the current baseline snapshots. They are
overwritten in place each time the owning test runs, so a `git diff` after a test run shows
exactly what changed on screen. Non-UI test projects (`UnitTests`, `IntegrationTests`,
`OpenXmlLibTests`, `SqliteDbTests._netF`, `RazorTests`) do not write here — they assert in
memory and produce no visual output.

To regenerate everything, run the full suite (`.\Scripts\Run-AllTests.ps1`) or a single project
(e.g. `dotnet test src\Tests\OqtaneTests\ExploreTheWorld.OqtaneTests.csproj`).

## Folder layout

Every UI test writes to a folder derived from its project and its fully-qualified test name:

```
TestResults/<Project>/<Category>/<TestClass>/<TestMethod>/<artifacts>
```

The path is built by `UITestSettings.GetTestFolder(...)` in each test project: `<Project>` is the
assembly name with the `JBC.ExploreTheWorld.` prefix stripped (so `._netF` variants keep their
suffix), and the remaining segments come from splitting the test's namespace-qualified display
name on `.`. `OutputRoot` resolves to this `TestResults/` directory at the repo root.

| Folder | Test project | Driver | Target app under test |
|--------|--------------|--------|-----------------------|
| `WebAppTests/` | `WebAppTests` | Playwright | `AL.BlazorWebApp` (hybrid Blazor web app) |
| `OqtaneTests/` | `OqtaneTests` | Playwright | ETW Oqtane modules + theme on the local Oqtane host |
| `OfficeWebAddinTests/` | `OfficeWebAddinTests` | Playwright | Word/Excel/PowerPoint Office.js Blazor web add-ins |
| `WinFormAppTests/` | `WinFormAppTests` | FlaUI | `AL.WinFormApp` (net10 BlazorWebView host) |
| `WinFormAppTests._netF/` | `WinFormAppTests._netF` | FlaUI | `AL.WinFormApp._netF` (net481 WinForms host) |
| `MauiAppTests/` | `MauiAppTests` | FlaUI | `AL.MauiApp.WinUI` (unpackaged MAUI Blazor exe) |
| `OfficeAddinTests/` | `OfficeAddinTests` | FlaUI | net10 VSTO COM add-in (ribbon + floating WebView form) |
| `OfficeAddinTests._netF/` | `OfficeAddinTests._netF` | FlaUI | net481 VSTO COM add-in |
| `AccessDbTests/` | `AccessDbTests` | FlaUI | `VBA/Access/ExploreTheWorld.accdb` in MSACCESS.EXE (read-only) |
| `_Exports/` | — (not a test project) | — | Sample Office documents from the Countries API export feature |

## Artifacts inside each test folder

| File | Written by | Meaning |
|------|-----------|---------|
| `before.png` | all UI tests | The app/page after it has launched and painted, captured **before** the action under test. Bases foreground the window and wait `ETW_TEST_SETTLE_MS` first, because WebView2/Blazor paints asynchronously and an immediate capture is blank. |
| `after.png` | all UI tests | The final state, captured at the end of the test (in teardown, best-effort). |
| `recording.webm` | Playwright projects only | Video of the whole browser session (`WebAppTests`, `OqtaneTests`, `OfficeWebAddinTests`). Playwright records to an auto-named file that the base renames to `recording.webm`. |
| other `*.png` (e.g. `form-open.png`, `countries-form-open.png`) | specific tests | Intermediate captures a test takes at a key step between `before` and `after`. |
| stray `page@<hash>.webm` | Playwright | The raw video left behind when the rename to `recording.webm` didn't run (the test aborted before teardown). Safe to delete; the paired `recording.webm` is the canonical copy. |

FlaUI projects produce PNGs only (no video). Screenshots and video can be toggled off, and the
capture behavior tuned, via environment variables read in each project's `UITestSettings`:

| Variable | Default | Effect |
|----------|---------|--------|
| `ETW_TEST_SCREENSHOTS` | on | Set to `false` to skip `before.png`/`after.png`. |
| `ETW_TEST_VIDEO` | on | Set to `false` to skip `recording.webm` (Playwright projects). |
| `ETW_TEST_HEADLESS` | `false` (visible) | Set to `true` to run the browser headless. |
| `ETW_TEST_SETTLE_MS` | 1500–4000 (per project) | Milliseconds to wait for the page/window to finish painting before capturing. |

## `_Exports/`

Curated **sample outputs of the Countries API "Export" feature** — the countries list rendered
into a Word/Excel/PowerPoint document once per Office-automation library, so the results of every
export path can be compared side-by-side:

```
ETW_CountriesNow-<Method>.{docx|xlsx|pptx}
```

The `-<Method>` token comes from `CL.MsOfficeExportName_Helper.BuildFileName`:

- `-Interop`, `-Dynamic`, `-NetOffice`, `-OpenXML` — the four desktop export libraries
  (`AL.WinFormApp` / `AL.MauiApp.WinUI` / VSTO add-ins / `AL.ExportData.ConsoleApp`).
- `-Web` — the browser OpenXML export (`BrowserExport_AppService`, used by the web app and Oqtane).
- `-Access` (where present) — the same document exported while the **AccessDb** provider is the
  active data source.

These are reference artifacts collected by hand from the running apps; they are not written by
`Run-AllTests.ps1`. They exist so a change to any writer can be diffed against a known-good file.
