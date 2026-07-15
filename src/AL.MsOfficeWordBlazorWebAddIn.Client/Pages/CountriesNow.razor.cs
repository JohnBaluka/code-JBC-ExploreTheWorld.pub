using System.Net.Http;
using System.Net.Http.Json;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.Json.Serialization;

using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl;

using Microsoft.AspNetCore.Components;

namespace ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class CountriesNow : ComponentBase
{
    [Inject]
    private HttpClient Http { get; set; } = default!;

    [Inject]
    private WordCountriesExport__Interop__Interface ExportInterop { get; set; } = default!;

    private CountryRow[]? _countries;
    private bool _isLoading;
    private string? _error;

    private bool _includeFlags = true;
    private bool _isExporting;
    private string? _exportStatus;
    private string? _exportError;

    // Flag PNG-thumbnail URLs (ISO2 → Wikimedia URL) are fetched once and reused.
    private static Dictionary<string, string>? _flagUrlCache;

    private async Task LoadCountriesAsync()
    {
        _isLoading    = true;
        _error        = null;
        _countries    = null;
        _exportStatus = null;
        _exportError  = null;
        StateHasChanged();
        try
        {
            using var resp = await Http.GetAsync("https://countriesnow.space/api/v0.1/countries/");
            resp.EnsureSuccessStatusCode();
            var body = await resp.Content.ReadFromJsonAsync<CountriesNowResponse>();
            _countries = body?.Data ?? [];
        }
        catch (Exception ex) { _error = ex.Message; }
        _isLoading = false;
        StateHasChanged();
    }

    private async Task ExportToDocumentAsync()
    {
        if (_countries is null || _countries.Length == 0) return;
        _isExporting  = true;
        _exportStatus = null;
        _exportError  = null;
        StateHasChanged();
        try
        {
            var flagUrls = _includeFlags
                ? await GetFlagUrlMapAsync()
                : new Dictionary<string, string>();

            var payload = new ExportPayload(_countries.Select(c => new ExportCountry(
                c.Country,
                c.Iso2,
                c.Iso3,
                flagUrls.TryGetValue(c.Iso2 ?? string.Empty, out var url) ? url : string.Empty)).ToList());

            var json   = JsonSerializer.Serialize(payload);
            var result = await ExportInterop.InsertCountriesAsync(json);

            if (!string.IsNullOrEmpty(result.Error))
                _exportError = result.Error;
            else
                _exportStatus = $"Inserted {result.Count} countries ({result.FlagCount} flags) into the document.";
        }
        catch (Exception ex) { _exportError = ex.Message; }
        _isExporting = false;
        StateHasChanged();
    }

    // Fetches the CountriesNow flag list and derives a Wikimedia PNG-thumbnail URL per ISO2.
    // Best-effort: on failure the export proceeds without flags.
    private async Task<Dictionary<string, string>> GetFlagUrlMapAsync()
    {
        if (_flagUrlCache is not null) return _flagUrlCache;
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var resp = await Http.GetFromJsonAsync<FlagsResponse>(
                "https://countriesnow.space/api/v0.1/countries/flag/images");
            foreach (var f in resp?.Data ?? [])
            {
                if (string.IsNullOrWhiteSpace(f.Iso2)) continue;
                var png = FlagImageUrl_Helper.GetPngThumbnailUrl(f.Flag);
                if (png is not null) map[f.Iso2!] = png;
            }
        }
        catch { /* flags are optional */ }
        _flagUrlCache = map;
        return map;
    }

    private sealed record CountryRow(string Country, string Iso2, string Iso3);

    private sealed class CountriesNowResponse
    {
        [JsonPropertyName("data")]
        public CountryRow[]? Data { get; set; }
    }

    // Payload passed to the Office.js interop (property names match countries-export.js).
    private sealed record ExportCountry(string country, string iso2, string iso3, string flagUrl);
    private sealed record ExportPayload(List<ExportCountry> countries);

    private sealed class FlagsResponse
    {
        [JsonPropertyName("data")]
        public FlagItem[]? Data { get; set; }
    }

    private sealed record FlagItem(
        [property: JsonPropertyName("iso2")] string? Iso2,
        [property: JsonPropertyName("flag")] string? Flag);
}
