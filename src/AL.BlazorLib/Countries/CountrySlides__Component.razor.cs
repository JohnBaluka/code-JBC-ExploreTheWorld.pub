using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.AL.BlazorLib.Countries
{
    public partial class CountrySlides__Component : Base__RadzenComponent, IAsyncDisposable
    {
        [Inject] protected CountriesNowSpaceManager__Service Service { get; set; } = default!;
        [Inject] protected RevealJs__Interop__Interface SlidesInterop { get; set; } = default!;

        // Optional — hosts without a registered flag image manager fall back to emoji flags.
        [Inject] protected FlagImageManager__Service? FlagImages { get; set; }

        // ── Data ────────────────────────────────────────────────────────────
        private List<CountryBasic_Row> _countries = new();
        private bool _isLoading = true;
        private string? _errorMessage;

        // ── Flag images (ISO2 → data URL, filled in the background) ─────────
        private readonly Dictionary<string, string> _flagImageSources = new(StringComparer.OrdinalIgnoreCase);
        private CancellationTokenSource? _flagLoadCts;

        // ── Reveal element refs ──────────────────────────────────────────────
        private ElementReference _revealEl;
        private ElementReference _wrapperEl;
        private bool _revealInitialized;
        private DotNetObjectReference<CountrySlides__Component>? _dotNetRef;

        // ── Slide state (updated via JS callbacks) ──────────────────────────
        private int _slideH;
        private int _slideV;
        private int _totalSlides;
        private bool _isOverview;
        private bool _isFullscreen;

        // ── Controls ────────────────────────────────────────────────────────
        private string _selectedTheme = "black";
        private string _selectedTransition = "slide";
        private bool _isAutoPlaying;
        private int _autoPlayInterval = 4000;
        private bool _showSlideNumbers = true;
        private string _notesPosition = "hidden"; // "hidden" | "right" | "bottom"
        private bool _searchVisible;

        protected static readonly List<ThemeOption_Row> Themes =
        [
            new("Black",     "black"),
            new("White",     "white"),
            new("League",    "league"),
            new("Beige",     "beige"),
            new("Sky",       "sky"),
            new("Night",     "night"),
            new("Serif",     "serif"),
            new("Simple",    "simple"),
            new("Solarized", "solarized"),
            new("Blood",     "blood"),
            new("Moon",      "moon"),
            new("Dracula",   "dracula"),
        ];

        protected static readonly List<string> Transitions =
        [
            "none", "slide", "convex", "concave", "zoom", "fade"
        ];

        protected static readonly List<AutoPlaySpeed_Row> AutoPlaySpeeds =
        [
            new("2 sec",  2000),
            new("4 sec",  4000),
            new("6 sec",  6000),
            new("10 sec", 10000),
        ];

        // True only in an interactive host with JavaScript (Server / WebAssembly / BlazorWebView).
        // Radzen Blazor Studio's design-time preview and static SSR pre-render report false —
        // RendererInfo throws "No renderer has been initialized" in the Studio preview — so the
        // reveal.js deck (which needs JS to lay out and inject its CSS) is swapped for a static
        // title-slide placeholder instead of collapsing to an empty black area.
        private bool IsInteractiveHost
        {
            get
            {
                try { return RendererInfo.IsInteractive; }
                catch (InvalidOperationException) { return false; }
            }
        }

        // ── Lifecycle ────────────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            try
            {
                var result = await Service.GetAllCountriesAsync();
                _countries = result.Data;
            }
            catch (Exception ex)
            {
                _errorMessage = $"Failed to load countries: {ex.Message}";
            }
            finally
            {
                _isLoading = false;
            }
        }

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            // In the design-time preview / static SSR the reveal.js deck isn't rendered
            // (see IsInteractiveHost) and there is no JS to drive it, so skip initialisation —
            // otherwise InitializeAsync fails on the missing element and surfaces a spurious
            // "Reveal.js failed to initialise" error, hiding the toolbar and placeholder.
            if (!IsInteractiveHost || _revealInitialized || _isLoading || _countries.Count == 0 || _errorMessage != null)
                return;

            _revealInitialized = true;
            _dotNetRef = DotNetObjectReference.Create(this);

            var ok = await SlidesInterop.InitializeAsync(_revealEl, _dotNetRef, _selectedTransition, _selectedTheme, 3);
            if (!ok)
            {
                _errorMessage = "Reveal.js failed to initialise. Check the browser console for details.";
                StateHasChanged();
                return;
            }

            // Load flag images in the background after the slideshow is interactive;
            // slides show emoji flags until the real image arrives.
            if (FlagImages != null)
            {
                _flagLoadCts = new CancellationTokenSource();
                _ = LoadFlagImagesAsync(_flagLoadCts.Token);
            }
        }

        // ── Flag image loading ───────────────────────────────────────────────
        private async Task LoadFlagImagesAsync(CancellationToken cancellationToken)
        {
            var loadedSinceRender = 0;
            foreach (var country in _countries)
            {
                if (cancellationToken.IsCancellationRequested) return;
                if (string.IsNullOrEmpty(country.Iso2) || _flagImageSources.ContainsKey(country.Iso2!))
                    continue;

                try
                {
                    var image = await FlagImages!.GetFlagImageAsync(country.Iso2!);
                    if (image != null)
                    {
                        _flagImageSources[image.Iso2] =
                            $"data:image/png;base64,{Convert.ToBase64String(image.ImageBytes)}";
                        loadedSinceRender++;
                    }
                }
                catch
                {
                    // Unavailable image (offline, 404, …) — keep the emoji fallback for this country.
                }

                if (loadedSinceRender >= 20)
                {
                    loadedSinceRender = 0;
                    await InvokeAsync(StateHasChanged);
                }
            }

            if (loadedSinceRender > 0 && !cancellationToken.IsCancellationRequested)
                await InvokeAsync(StateHasChanged);
        }

        private string? GetFlagImageSrc(string? iso2) =>
            iso2 != null && _flagImageSources.TryGetValue(iso2, out var src) ? src : null;

        // ── JS callbacks ─────────────────────────────────────────────────────
        [JSInvokable]
        public void OnSlideChanged(int h, int v, int total)
        {
            _slideH = h;
            _slideV = v;
            _totalSlides = total;
            if (_isAutoPlaying && h == _totalSlides - 1)
                _isAutoPlaying = false;
            StateHasChanged();
        }

        [JSInvokable]
        public void OnOverviewChanged(bool isOverview)
        {
            _isOverview = isOverview;
            StateHasChanged();
        }

        [JSInvokable]
        public void OnFullscreenChanged(bool isFullscreen)
        {
            _isFullscreen = isFullscreen;
            StateHasChanged();
        }

        // ── Navigation ───────────────────────────────────────────────────────
        private async Task NavNextAsync() => await SlidesInterop.NavigateNextAsync();
        private async Task NavPrevAsync() => await SlidesInterop.NavigatePrevAsync();
        private async Task NavToAsync(int h) => await SlidesInterop.NavigateToSlideAsync(h, 0);
        private async Task ToggleOverviewAsync() => await SlidesInterop.ToggleOverviewAsync();

        // ── Theme / Transition ───────────────────────────────────────────────
        private async Task OnThemeChangedAsync(object? value)
        {
            if (value is string theme)
            {
                _selectedTheme = theme;
                await SlidesInterop.SetThemeAsync(theme);
            }
        }

        private async Task OnTransitionChangedAsync(object? value)
        {
            if (value is string t)
            {
                _selectedTransition = t;
                await SlidesInterop.SetTransitionAsync(t);
            }
        }

        private async Task ToggleSlideNumbersAsync()
        {
            _showSlideNumbers = !_showSlideNumbers;
            await SlidesInterop.SetSlideNumberAsync(_showSlideNumbers);
        }

        // ── Auto-play ────────────────────────────────────────────────────────
        private async Task ToggleAutoPlayAsync()
        {
            _isAutoPlaying = !_isAutoPlaying;
            if (_isAutoPlaying)
                await SlidesInterop.StartAutoPlayAsync(_autoPlayInterval);
            else
                await SlidesInterop.StopAutoPlayAsync();
        }

        private async Task OnAutoPlaySpeedChangedAsync(object? value)
        {
            if (value is int ms)
            {
                _autoPlayInterval = ms;
                if (_isAutoPlaying)
                {
                    await SlidesInterop.StopAutoPlayAsync();
                    await SlidesInterop.StartAutoPlayAsync(ms);
                }
            }
        }

        // ── Notes position ────────────────────────────────────────────────────
        private async Task CycleNotesPositionAsync()
        {
            _notesPosition = _notesPosition switch
            {
                "hidden" => "right",
                "right"  => "bottom",
                _        => "hidden"
            };
            await SlidesInterop.SetNotesPositionAsync(_notesPosition);
        }

        private string NotesIcon => _notesPosition switch
        {
            "right"  => "speaker_notes",
            "bottom" => "vertical_split",
            _        => "speaker_notes_off"
        };

        private string NotesTitle => _notesPosition switch
        {
            "hidden" => "Show notes (right)",
            "right"  => "Move notes to bottom",
            _        => "Hide notes"
        };

        // ── Search ────────────────────────────────────────────────────────────
        private async Task ToggleSearchAsync()
        {
            _searchVisible = !_searchVisible;
            await SlidesInterop.ToggleSearchAsync();
        }

        // ── Fullscreen ────────────────────────────────────────────────────────
        private async Task ToggleFullscreenAsync() => await SlidesInterop.RequestFullscreenAsync(_wrapperEl);

        // ── Helpers ───────────────────────────────────────────────────────────
        private double SlideProgress =>
            _totalSlides > 1 ? (_slideH / (double)(_totalSlides - 1)) * 100 : 0;

        private string SlideLabel =>
            _totalSlides > 0 ? $"{_slideH + 1} / {_totalSlides}" : "—";

        private static string GetFlagEmoji(string? iso2)
        {
            if (iso2?.Length != 2) return "🌐";
            return string.Concat(
                iso2.ToUpperInvariant().Select(c => char.ConvertFromUtf32(0x1F1E6 + (c - 'A'))));
        }

        private static string GetSlideBackground(string? iso2)
        {
            if (string.IsNullOrEmpty(iso2)) return "#1a1a2e";
            var hash = iso2.Aggregate(0, (h, c) => h * 31 + c);
            var hue = Math.Abs(hash) % 360;
            return $"hsl({hue}, 50%, 20%)";
        }

        // ── Dispose ───────────────────────────────────────────────────────────
        public async ValueTask DisposeAsync()
        {
            _flagLoadCts?.Cancel();
            _flagLoadCts?.Dispose();

            if (_isAutoPlaying)
                try { await SlidesInterop.StopAutoPlayAsync(); } catch { }

            try { await SlidesInterop.DestroyAsync(); } catch { }

            _dotNetRef?.Dispose();
        }
    }
}
