using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;
using Microsoft.AspNetCore.Components;
using Radzen.Blazor;

namespace JBC.ExploreTheWorld.AL.BlazorLib.Countries
{
    public partial class CountriesNow__Component : Base__RadzenComponent
    {
        [Inject] protected CountriesNowSpaceManager__Service Service { get; set; } = default!;
        [Inject] protected OfficeExport_AppService__Interface? Export_AppService { get; set; }
        [Inject] private DbProvider_AppService DbProvider_AppService { get; set; } = default!;
        [Inject] private Layout_AppService Layout_AppService { get; set; } = default!;

        private List<CountryBasic_Row> _countries = new();
        private IList<CountryBasic_Row>? _selectedCountries;
        private CountryBasic_Row? _selectedCountry;

        private List<CountryState_Row> _states = new();

        private bool _isLoading = true;
        private bool _isLoadingStates;
        private string? _errorMessage;
        private string? _statesErrorMessage;

        private DataSource_Enum? _countriesSource;
        private DataSource_Enum? _statesSource;

        private RadzenDataGrid<CountryBasic_Row> _countriesGrid = default!;

        private readonly List<string> _exportTypes     = new() { "Word", "Excel", "PowerPoint" };
        // Populated from the host's export service in OnInitialized: browser hosts offer OpenXML
        // only, desktop hosts offer Interop/Dynamic/NetOffice/OpenXML.
        private List<string> _exportLibraries = new() { "OpenXML" };
        private string _exportType     = "Word";
        private string _exportLibrary  = "OpenXML";
        private string _exportFilePath = string.Empty;
        private string _exportLog      = string.Empty;

        // Desktop hosts write to a local path; browser hosts stream a download by file name only.
        private bool SavesToFilePath => Export_AppService?.SavesToFilePath ?? true;
        private string ExportPathLabel => SavesToFilePath ? "File Path" : "File Name";
        private string ExportPathPlaceholder => SavesToFilePath
            ? @"e.g. C:\Users\...\Documents\export.docx"
            : "e.g. ETW_CountriesNow.docx";

        protected override void OnInitialized()
        {
            // The library picklist reflects what this host's export service actually supports.
            if (Export_AppService is { SupportedLibraries.Count: > 0 } service)
            {
                _exportLibraries = new List<string>(service.SupportedLibraries);
                if (!_exportLibraries.Contains(_exportLibrary))
                    _exportLibrary = _exportLibraries[0];
            }

            UpdateDefaultExportPath();
            DbProvider_AppService.OnProviderChanged += HandleProviderChanged;
        }

        // Dispatched on the Blazor sync context so StateHasChanged is safe to call inside.
        private void HandleProviderChanged() => InvokeAsync(async () =>
        {
            RefreshDefaultExportPathIfDefault();
            await ReloadCountriesAsync();
        });

        protected override async Task OnInitializedAsync()
        {
            await ReloadCountriesAsync();
        }

        private async Task ReloadCountriesAsync()
        {
            _isLoading       = true;
            _countries       = new();
            _states          = new();
            _selectedCountry = null;
            _selectedCountries = null;
            _countriesSource = null;
            _statesSource    = null;
            _errorMessage    = null;
            _statesErrorMessage = null;
            _exportLog       = string.Empty;
            StateHasChanged();

            AppendExportLog($"Loading countries ({DbProvider_AppService.ProviderName})...");
            try
            {
                var result = await Service.GetAllCountriesAsync();
                _countries       = result.Data;
                _countriesSource = result.Source;
                AppendExportLog($"✔ Loaded {_countries.Count} countries (source: {result.Source}).");
                AppendExportLog("Ready. Select a country or click Export.");
            }
            catch (Exception ex)
            {
                _errorMessage = $"Failed to load countries: {ex.Message}";
                AppendExportLog($"✘ Failed to load countries: {ex.Message}");
            }
            finally
            {
                _isLoading = false;
                StateHasChanged();
            }
        }

        private void UpdateDefaultExportPath()
        {
            var ext = _exportType switch
            {
                "Excel"      => "xlsx",
                "PowerPoint" => "pptx",
                _            => "docx"
            };
            var fromAccessDb = string.Equals(
                DbProvider_AppService.ProviderName, DbProviderNames.AccessDb, StringComparison.OrdinalIgnoreCase);
            var fileName = MsOfficeExportName_Helper.BuildFileName(
                "ETW_CountriesNow", _exportLibrary, fromAccessDb, ext);

            // Browser hosts download by file name only (no local file system) — default to the
            // bare file name so the field never implies a local save path. Desktop hosts default
            // to the full path in the user's Documents folder.
            if (!SavesToFilePath)
            {
                _exportFilePath = fileName;
                return;
            }

            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            _exportFilePath = string.IsNullOrEmpty(documents)
                ? fileName
                : Path.Combine(documents, fileName);
        }

        // When the export type / library / DB provider changes, refresh the default export file
        // name (extension, method token, Access token), unless the user typed a custom path.
        private void OnExportTypeChanged(object value) => RefreshDefaultExportPathIfDefault();

        private void OnExportLibraryChanged(object value) => RefreshDefaultExportPathIfDefault();

        private void RefreshDefaultExportPathIfDefault()
        {
            // Compare the file name (not the whole path) so our generated Documents-folder default
            // is still recognized and refreshed, while a user-typed custom path is preserved.
            var fileName = string.IsNullOrEmpty(_exportFilePath)
                ? string.Empty
                : Path.GetFileName(_exportFilePath);
            if (string.IsNullOrEmpty(_exportFilePath) ||
                fileName.StartsWith("ETW_CountriesNow", StringComparison.OrdinalIgnoreCase))
            {
                UpdateDefaultExportPath();
            }
        }

        private async Task OnCountrySelected(CountryBasic_Row country)
        {
            _selectedCountry    = country;
            _states             = new();
            _statesSource       = null;
            _statesErrorMessage = null;
            _isLoadingStates    = true;
            AppendExportLog($"Loading states for {country.Country}...");
            StateHasChanged();

            try
            {
                var result    = await Service.GetCountryStatesAsync(country.Country ?? "");
                _states       = result.Data;
                _statesSource = result.Source;
                AppendExportLog($"✔ Loaded {_states.Count} states for {country.Country} (source: {result.Source}).");
            }
            catch (Exception ex)
            {
                _statesErrorMessage = $"Failed to load states: {ex.Message}";
                AppendExportLog($"✘ Failed to load states for {country.Country}: {ex.Message}");
            }
            finally
            {
                _isLoadingStates = false;
            }
        }

        private async Task ClearDatabaseAsync()
        {
            try
            {
                await Service.ClearAllDataAsync();
                _countries         = new();
                _states            = new();
                _countriesSource   = null;
                _statesSource      = null;
                _selectedCountry   = null;
                _selectedCountries = null;
                _isLoading         = true;
                StateHasChanged();

                var result = await Service.GetAllCountriesAsync();
                _countries       = result.Data;
                _countriesSource = result.Source;
            }
            catch (Exception ex)
            {
                _errorMessage = $"Failed to clear database: {ex.Message}";
            }
            finally
            {
                _isLoading = false;
            }
        }

        private void ClearExportLog() => _exportLog = string.Empty;

        private async Task ExportAsync()
        {
            if (string.IsNullOrWhiteSpace(_exportFilePath))
            {
                AppendExportLog("ERROR: No export file path specified.");
                return;
            }

            if (Export_AppService == null)
            {
                AppendExportLog("ERROR: Export service not available in this host.");
                return;
            }

            AppendExportLog($"Export type    : {_exportType}");
            AppendExportLog($"Export library : {_exportLibrary}");
            AppendExportLog(SavesToFilePath
                ? $"Target         : {_exportFilePath}"
                : $"Download       : {_exportFilePath}");
            AppendExportLog($"Countries      : {_countries.Count}");

            try
            {
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                await Export_AppService.ExportCountriesAsync(
                    _countries, _exportType, _exportLibrary, _exportFilePath,
                    msg => AppendExportLog(msg));
                stopwatch.Stop();
                AppendExportLog($"✔ Export completed in {Duration_Helper.Format(stopwatch.Elapsed)}.");
            }
            catch (Exception ex)
            {
                AppendExportLog($"✘ Export failed: {ex.Message}");
            }
        }

        private void AppendExportLog(string message)
        {
            _exportLog += $"[{DateTime.Now:HH:mm:ss.fff}] {message}\n";
        }

        public override void Dispose()
        {
            DbProvider_AppService.OnProviderChanged -= HandleProviderChanged;
            base.Dispose();
        }
    }
}
