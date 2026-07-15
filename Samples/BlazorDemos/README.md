# BlazorDemos

This solution demonstrates three ways to build the same Blazor component: enter a
name, click the button, and see a greeting pop up. The demos are functionally
equivalent but differ in how the markup, code, styling, and "alert" are provided.

## Projects

- `BlazorDemos` - Blazor Web App (server host + prerendering).
- `BlazorDemos.Client` - WebAssembly project that contains the interactive pages.

The pages use `@rendermode InteractiveAuto`: the first visit is handled by the
server over SignalR while the WebAssembly runtime downloads in the background,
then subsequent interactions run in the browser via WebAssembly.

## Demos

- `/all-in-one-demo` — everything in a single file
  - `AllInOneDemo.razor`
  - Razor markup, C# code, component CSS (`<style>` block), and the JavaScript
    (`<script>` block) all live in this one file — **no separate `.js` file is
    needed**.
  - Because the page is prerendered, the inline `<script>` is part of the initial
    HTML and runs in the browser (Blazor also re-runs it during enhanced
    navigation), registering the greeting function before the button is clicked.

- `/separate-files-demo` — the conventional split-file layout
  - `SeparateFilesDemo.razor` — markup
  - `SeparateFilesDemo.razor.cs` — code-behind (partial class)
  - `SeparateFilesDemo.razor.css` — scoped CSS
  - `SeparateFilesDemo.razor.js` — collocated JavaScript module (served
    automatically at `./Components/Pages/SeparateFilesDemo.razor.js`)

- `/radzen-demo` — component library, no CSS or JS files
  - `RadzenDemo.razor` — markup (Radzen components only)
  - `RadzenDemo.razor.cs` — code-behind (partial class)
  - Styling comes entirely from the Radzen theme, so **no `.razor.css` is needed**.
  - The greeting is shown with Radzen's `DialogService.Alert(...)`, a dialog
    rendered by Blazor — so **no `.razor.js` / JS interop is needed**.
  - `<RadzenComponents />` (the dialog/notification host) is placed on the page
    itself so it shares the page's interactive DI scope, and therefore the same
    `DialogService` instance the code-behind injects.

## Run

```bash
dotnet run --project BlazorDemos/BlazorDemos.csproj
```

Then open the displayed local URL and navigate to any demo page.

## Notes

- The projects target `net10.0`. To run on .NET 9, change the target framework
  in both `BlazorDemos.csproj` and `BlazorDemos.Client.csproj` from `net10.0`
  to `net9.0`.
- The host serves framework and static assets with `app.MapStaticAssets()`
  (the .NET 9+ replacement for `app.UseStaticFiles()`). This is what serves
  `_framework/blazor.web.js`; without it the pages render but never become
  interactive, so the buttons do nothing.
- Scoped component CSS from the WebAssembly project is bundled automatically
  into `BlazorDemos.styles.css` — there is no need to reference the per-project
  `*.bundle.scp.css` file directly.
- The all-in-one demo relies on its inline `<script>` executing during
  prerendering / enhanced navigation. A strict `script-src` Content Security
  Policy would need to permit inline scripts (the separate-files collocated
  `.razor.js` module approach avoids that consideration).
- The Radzen demo uses [Radzen.Blazor](https://blazor.radzen.com/). Setup lives
  in the host: `AddRadzenComponents()` is called in both `Program.cs` files, and
  `App.razor` references the Radzen theme (`_content/Radzen.Blazor/css/default.css`)
  and script (`_content/Radzen.Blazor/Radzen.Blazor.js`). That script is Radzen
  library infrastructure, not demo-authored interop.
