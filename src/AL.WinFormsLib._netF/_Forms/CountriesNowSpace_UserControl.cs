using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
#if !NETFRAMEWORK
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using Microsoft.Extensions.DependencyInjection;
#endif

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF0123456701")]
    public partial class CountriesNowSpace_UserControl : UserControl
#if !NETFRAMEWORK
        , ICustomQueryInterface, IObjectSafety
#endif
    {
        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x1;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x2;
#if !NETFRAMEWORK
        // Hides IProvideClassInfo / IProvideClassInfo2 from the Office CTP host so that
        // .NET 10 UserControls can be hosted as VSTO task panes without aborting on
        // missing typelib metadata (HRESULT 0x80131165).
        private static readonly Guid _iidIProvideClassInfo  = new Guid("B196B283-BAB4-101A-B69C-00AA00341D07");
        private static readonly Guid _iidIProvideClassInfo2 = new Guid("A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851");
        CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            if (iid == _iidIProvideClassInfo || iid == _iidIProvideClassInfo2)
                return CustomQueryInterfaceResult.Failed;
            return CustomQueryInterfaceResult.NotHandled;
        }

        int IObjectSafety.GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions)
        {
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            return 0;
        }

        int IObjectSafety.SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions) => 0;
#endif
        private readonly CountriesNowSpaceManager__Service _manager;
        private readonly MsOfficeExportManager__Service _exportManager;
#if !NETFRAMEWORK
        private ServiceProvider? _ownedServices; // non-null only when we built our own DI
#endif
        private string? _selectedCountry = null;

        // Add-in export mode: set via EnableAddinExportMode when the control is hosted inside a
        // VSTO add-in. Export type is locked to the host Office app, the library is locked to
        // NetOffice, and Export creates a new document in the running host application.
        private bool   _addinExportMode;
        private string _addinHostType = "Word";

        // Parameterless ctor — used by COM/CTP factory when instantiated as a task pane.
        // Builds its own ServiceProvider from default local path.
        // Only available on net10.0 (comhost task pane registration is net10.0-only).
#if !NETFRAMEWORK
        public CountriesNowSpace_UserControl()
        {
            var dbPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "JBC", "ExploreTheWorld", "etw.db");
            Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);
            var sc = new ServiceCollection();
            sc.AddTransient<CountriesNowSpaceApi_Interface, CountriesNowSpaceApi__Repo>();
            sc.AddExploreTheWorldSqliteDb($"Data Source={dbPath}");
            sc.AddTransient<CountriesNowSpaceManager__Service>();
            _ownedServices = sc.BuildServiceProvider();
            try { _ownedServices.EnsureExploreTheWorldDbCreated(); } catch { }
            _manager       = _ownedServices.GetRequiredService<CountriesNowSpaceManager__Service>();
            _exportManager = CreateExportManager(_manager);

            InitializeComponent();
            InitComboBoxes();
        }
#endif

        public CountriesNowSpace_UserControl(CountriesNowSpaceManager__Service manager)
        {
            InitializeComponent();
            _manager       = manager;
            _exportManager = CreateExportManager(manager);
            InitComboBoxes();
        }

        // Flag images are resolved via the %LocalAppData% file cache with Wikimedia download
        // fallback, so exports embed flags even when this control builds its own services.
        private static MsOfficeExportManager__Service CreateExportManager(
            CountriesNowSpaceManager__Service countriesNowManager)
            => new MsOfficeExportManager__Service(
                MsOfficeExportRepoFactoryProvider.Require(),
                new FlagImageManager__Service(
                    new FlagImageStore_FileSystem__Repo(),
                    new FlagImageDownload__Repo(),
                    countriesNowManager));

        private void InitComboBoxes()
        {
            cmbExportType.Items.AddRange(new object[] { "Word", "Excel", "PowerPoint" });
            cmbExportType.SelectedIndex = 0;

            cmbExportLibrary.Items.AddRange(new object[] { "Interop", "Dynamic", "NetOffice", "OpenXML" });
            cmbExportLibrary.SelectedItem = "NetOffice";

            // Show the full default export path (Documents folder) on load, and keep it in sync when
            // the export type or library changes — unless the user has typed their own path.
            cmbExportType.SelectedIndexChanged    += (_, _) => UpdateDefaultExportPath();
            cmbExportLibrary.SelectedIndexChanged += (_, _) => UpdateDefaultExportPath();
            UpdateDefaultExportPath();
        }

        // Full default export path shown in the textbox; used to detect a user-customized path.
        private string _defaultExportPath = string.Empty;

        private void UpdateDefaultExportPath()
        {
            if (_addinExportMode) return;

            // Preserve a path the user typed or picked; only overwrite our own generated default.
            var current = txtExportFilePath.Text.Trim();
            if (!string.IsNullOrEmpty(current) && current != _defaultExportPath)
                return;

            var type      = cmbExportType.SelectedItem?.ToString() ?? "Word";
            var library   = cmbExportLibrary.SelectedItem?.ToString() ?? "OpenXML";
            var extension = type switch { "Excel" => "xlsx", "PowerPoint" => "pptx", _ => "docx" };
            var fileName  = MsOfficeExportName_Helper.BuildFileName(
                "ETW_CountriesNow", library, fromAccessDb: false, extension);

            _defaultExportPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), fileName);
            txtExportFilePath.Text = _defaultExportPath;
        }

        // ── Add-in export mode ─────────────────────────────────────────────────────────

        /// <summary>
        /// Switches the export UI into add-in mode. The export type is locked to
        /// <paramref name="hostType"/> ("Word"/"Excel"/"PowerPoint") — the Office application the
        /// add-in is hosted in — and the library is locked to NetOffice. Export then creates a new
        /// document in the running host application instead of writing a file, and every export-row
        /// control except the Export button is hidden.
        /// </summary>
        public void EnableAddinExportMode(string hostType)
        {
            _addinExportMode = true;
            _addinHostType   = string.IsNullOrWhiteSpace(hostType) ? "Word" : hostType;
            ApplyAddinExportMode();
        }

        private void ApplyAddinExportMode()
        {
            // Lock the type to the host app and the library to NetOffice.
            cmbExportType.SelectedItem    = _addinHostType;
            cmbExportLibrary.SelectedItem = "NetOffice";

            // Only the Export button stays visible in the export row.
            lblExportType.Visible     = false;
            cmbExportType.Visible     = false;
            lblExportLibrary.Visible  = false;
            cmbExportLibrary.Visible  = false;
            lblExportFilePath.Visible = false;
            txtExportFilePath.Visible = false;
            btnBrowseExport.Visible   = false;
            btnClearLog.Visible       = false;
        }

        // ── Source indicator helpers ──────────────────────────────────────────────────

        private static void SetSourceLabel(Label label, DataSource_Enum source)
        {
            label.Text      = $"Source: {source}";
            label.ForeColor = source == DataSource_Enum.Database ? Color.ForestGreen : Color.SteelBlue;
        }

        // ── Grid data helper ─────────────────────────────────────────────────────────

        private DataTable GetGridDataAsTable()
        {
            var dt = new DataTable("Countries Now");
            foreach (DataGridViewColumn col in dgvCountriesNow.Columns)
                dt.Columns.Add(col.HeaderText);
            foreach (DataGridViewRow row in dgvCountriesNow.Rows)
            {
                if (row.IsNewRow) continue;
                var dr = dt.NewRow();
                for (int i = 0; i < dgvCountriesNow.Columns.Count; i++)
                    dr[i] = row.Cells[i].Value?.ToString() ?? "";
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private bool EnsureDataLoaded()
        {
            if (dgvCountriesNow.DataSource == null)
            {
                MessageBox.Show("No data loaded. Please click \"Load Countries\" first.",
                    "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        // ── Load handlers ────────────────────────────────────────────────────────────

        private async void btnLoadCountriesNow_Click(object sender, EventArgs e)
        {
            btnLoadCountriesNow.Enabled = false;
            AppendExportLog("Loading countries...");
            try
            {
                var result = await _manager.GetAllCountriesAsync();
                dgvCountriesNow.DataSource = result.Data
                    .OrderBy(c => c.Country)
                    .Select(c => new
                    {
                        Country = c.Country,
                        ISO2    = c.Iso2,
                        ISO3    = c.Iso3
                    }).ToList();
                SetSourceLabel(lblCountriesSource, result.Source);
                AppendExportLog($"Loaded {result.Data.Count} countries (Source: {result.Source})");
            }
            catch (Exception ex)
            {
                AppendExportLog($"Error loading countries: {ex.Message}");
                MessageBox.Show($"Error loading countries: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnLoadCountriesNow.Enabled = true;
            }
        }

        private void dgvCountriesNow_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvCountriesNow.SelectedRows.Count > 0)
            {
                _selectedCountry = dgvCountriesNow.SelectedRows[0].Cells["Country"].Value?.ToString();
                lblSelectedCountry.Text = $"Selected: {_selectedCountry}";
                btnLoadStates.Enabled = !string.IsNullOrEmpty(_selectedCountry);
            }
        }

        private async void btnLoadStates_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedCountry))
                return;

            var country = _selectedCountry!;
            btnLoadStates.Enabled = false;
            try
            {
                var result = await _manager.GetCountryStatesAsync(country);
                dgvStates.DataSource = result.Data
                    .OrderBy(s => s.Name)
                    .Select(s => new
                    {
                        State = s.Name,
                        Code  = s.StateCode
                    }).ToList();
                SetSourceLabel(lblStatesSource, result.Source);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading states: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnLoadStates.Enabled = !string.IsNullOrEmpty(_selectedCountry);
            }
        }

        // ── Clear DB handler ─────────────────────────────────────────────────────────

        private async void btnClearDb_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show(
                "This will delete all cached CountriesNow.space data from the database. Continue?",
                "Clear Database", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) return;

            btnClearDb.Enabled = false;
            AppendExportLog("Clearing database...");
            try
            {
                await _manager.ClearAllDataAsync();
                dgvCountriesNow.DataSource = null;
                dgvStates.DataSource       = null;
                lblCountriesSource.Text    = string.Empty;
                lblStatesSource.Text       = string.Empty;
                lblSelectedCountry.Text    = "Select a country to load states";
                _selectedCountry           = null;
                btnLoadStates.Enabled      = false;
                AppendExportLog("Database cleared.");
                MessageBox.Show("Database cleared successfully.", "Cleared",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendExportLog($"Error clearing database: {ex.Message}");
                MessageBox.Show($"Error clearing database: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnClearDb.Enabled = true;
            }
        }

        // ── Export controls ───────────────────────────────────────────────────────────

        private void btnBrowseExport_Click(object sender, EventArgs e)
        {
            var type = cmbExportType.SelectedItem?.ToString() ?? "Word";

            // Seed the dialog from the current path when set, otherwise the Documents folder.
            var current   = txtExportFilePath.Text.Trim();
            var initialDir = !string.IsNullOrEmpty(current) && Directory.Exists(Path.GetDirectoryName(current))
                ? Path.GetDirectoryName(current)!
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var fileName = !string.IsNullOrEmpty(current)
                ? Path.GetFileName(current)
                : MsOfficeExportName_Helper.BuildStem(
                    "ETW_CountriesNow", cmbExportLibrary.SelectedItem?.ToString() ?? "OpenXML", fromAccessDb: false);

            using var dlg = new SaveFileDialog
            {
                Title = "Select Export File Location",
                InitialDirectory = initialDir,
                Filter = type switch
                {
                    "Excel"      => "Excel Workbook (*.xlsx)|*.xlsx",
                    "PowerPoint" => "PowerPoint Presentation (*.pptx)|*.pptx",
                    _            => "Word Document (*.docx)|*.docx"
                },
                DefaultExt = type switch
                {
                    "Excel"      => "xlsx",
                    "PowerPoint" => "pptx",
                    _            => "docx"
                },
                FileName = fileName
            };
            if (dlg.ShowDialog(this) == DialogResult.OK)
                txtExportFilePath.Text = dlg.FileName; // a picked path differs from the generated
                                                       // default, so it is preserved on later changes
        }

        private async void btnExport_Click(object sender, EventArgs e)
        {
            if (!EnsureDataLoaded()) return;

            if (_addinExportMode)
            {
                await RunAddinExportAsync();
                return;
            }

            var filePath = txtExportFilePath.Text.Trim();
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Please specify an export file path.", "No File Path",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var exportType    = cmbExportType.SelectedItem?.ToString() ?? "Word";
            var exportLibrary = cmbExportLibrary.SelectedItem?.ToString() ?? "OpenXML";

            rtbExportLog.Clear();
            btnExport.Enabled = false;

            try
            {
                var data = GetGridDataAsTable();
                AppendExportLog($"Export type    : {exportType}");
                AppendExportLog($"Export library : {exportLibrary}");
                AppendExportLog($"Target         : {filePath}");
                AppendExportLog($"Rows           : {data.Rows.Count}");
                AppendExportLog("─────────────────────────────────────────────────────");

                await RunExportAsync(exportType, exportLibrary, filePath, data);

                AppendExportLog("─────────────────────────────────────────────────────");
                AppendExportLog("✔ Export completed successfully.");

                // Open the appropriate Watcher form and the exported file
                OpenWatcherForm(exportType, filePath);
                OpenExportedFile(filePath);
            }
            catch (Exception ex)
            {
                AppendExportLog("─────────────────────────────────────────────────────");
                AppendExportLog($"✘ Export failed: {ex.Message}");
                MessageBox.Show(ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnExport.Enabled = true;
            }
        }

        private void btnClearLog_Click(object sender, EventArgs e)
        {
            rtbExportLog.Clear();
        }

        private void AppendExportLog(string message)
        {
            if (InvokeRequired)
            {
                Invoke(() => AppendExportLog(message));
                return;
            }
            rtbExportLog.AppendText($"[{DateTime.Now:HH:mm:ss.fff}] {message}{Environment.NewLine}");
            rtbExportLog.ScrollToCaret();
        }

        private static void OpenExportedFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                try { Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true }); }
                catch { /* best-effort */ }
            }
        }

        private static void OpenWatcherForm(string exportType, string filePath)
        {
#if !MSOFFICE_ADDIN
            Form? watcher = exportType switch
            {
                "Word"        => new MsWord_Watcher_Form(),
                "Excel"       => new MsExcel_Watcher_Form(),
                "PowerPoint"  => new MsPowerPoint_Watcher_Form(),
                _             => null
            };
            if (watcher == null) return;
            watcher.Show();
#endif
        }

        // ── Export dispatch ───────────────────────────────────────────────────────────

        private Task RunExportAsync(string exportType, string exportLibrary, string filePath, DataTable data)
        {
            // Map the grid rows to the DL row type and delegate to the BL export manager.
            // UI items never build documents inline nor instantiate DL repositories directly.
            return _exportManager.ExportAsync(MapRows(data), exportType, exportLibrary, filePath, AppendExportLog);
        }

        private static IList<MsOfficeCountry_Row> MapRows(DataTable data) =>
            data.Rows.Cast<DataRow>()
                .Select(r => new MsOfficeCountry_Row(
                    r["Country"]?.ToString() ?? string.Empty,
                    r["ISO2"]?.ToString() ?? string.Empty,
                    r["ISO3"]?.ToString() ?? string.Empty))
                .ToList();

        // Add-in export: create a new document in the running host Office application (NetOffice)
        // and fill it with the grid data — no file is written. The host application is acquired on
        // this UI/STA thread so the COM proxy is valid here; the export runs on this thread too.
        private async Task RunAddinExportAsync()
        {
            rtbExportLog.Clear();
            btnExport.Enabled = false;

            object? hostApp = null;
            try
            {
                var progId = _addinHostType switch
                {
                    "Excel"      => "Excel.Application",
                    "PowerPoint" => "PowerPoint.Application",
                    _            => "Word.Application"
                };

                AppendExportLog($"Connecting to running {_addinHostType}...");
                var raw = WatcherComHelper.GetActiveCom(progId);
                hostApp = _addinHostType switch
                {
                    "Excel"      => new NetOffice.ExcelApi.Application(null, raw),
                    "PowerPoint" => new NetOffice.PowerPointApi.Application(null, raw),
                    _            => (object)new NetOffice.WordApi.Application(null, raw)
                };

                var rows = MapRows(GetGridDataAsTable());
                AppendExportLog($"Export type    : {_addinHostType}");
                AppendExportLog("Export library : NetOffice (new document)");
                AppendExportLog($"Rows           : {rows.Count}");
                AppendExportLog("─────────────────────────────────────────────────────");

                await _exportManager.ExportToRunningAppAsync(_addinHostType, hostApp, rows, AppendExportLog);

                AppendExportLog("─────────────────────────────────────────────────────");
                AppendExportLog("✔ Export completed successfully.");
            }
            catch (Exception ex)
            {
                AppendExportLog("─────────────────────────────────────────────────────");
                AppendExportLog($"✘ Export failed: {ex.Message}");
                MessageBox.Show(ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Release our COM wrapper (does not quit the user's Office instance).
                (hostApp as IDisposable)?.Dispose();
                btnExport.Enabled = true;
            }
        }
    }
}
