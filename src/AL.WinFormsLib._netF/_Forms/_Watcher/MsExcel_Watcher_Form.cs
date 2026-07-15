using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using JBC.ExploreTheWorld.CL;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class MsExcel_Watcher_Form : Form
    {
        // ── State ────────────────────────────────────────────────────────────────────
        private readonly List<MsOfficeEvent_Record> _events = MsOfficeEvents_Repo.ForExcel();
        private ComboBox? _cbxDocuments;
        private bool _isConnected;
        private bool _applicationInjected;

        // Last auto-generated default output path — used so a user-entered custom path is preserved
        // when the active workbook changes.
        private string _lastDefaultOutputPath = string.Empty;

        // NetOffice
        private NetOffice.ExcelApi.Application? _excelAppNo;
        private NetOffice.ExcelApi.Workbook?    _excelWbNo;

        // ── Constructor ──────────────────────────────────────────────────────────────
        public MsExcel_Watcher_Form()
        {
            InitializeComponent();
            _cbxDocuments = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 260,
                Location = new System.Drawing.Point(12, 10),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _cbxDocuments.SelectedIndexChanged += cbxDocuments_SelectedIndexChanged;
            cbxJsonWriteMethod.SelectedIndexChanged += cbxJsonWriteMethod_SelectedIndexChanged;
            pnlHeader.Controls.Add(_cbxDocuments);
            pnlHeader.Controls.SetChildIndex(_cbxDocuments, 0);
        }

        // ── Load ─────────────────────────────────────────────────────────────────────
        private void MsExcel_Watcher_Form_Load(object sender, EventArgs e)
        {
            _lastDefaultOutputPath = SaveAsJson_Helper.BuildDefaultPath(TrySafeWbFullName(), "ETW_MsExcel", cbxJsonWriteMethod.SelectedItem?.ToString());
            txtOutputFilePath.Text = _lastDefaultOutputPath;
            cbxJsonWriteMethod.Items.AddRange(Enum.GetNames(typeof(JsonWriteMethod)));
            cbxJsonWriteMethod.SelectedItem = JsonWriteMethod.NetOffice.ToString();
            PopulateEventsGrid();
            if (_applicationInjected)
            {
                btnConnectDisconnect.Text    = "Disconnect";
                btnConnectDisconnect.Enabled = false;
                AppendLog($"Connected (via addin): {_excelWbNo?.Name ?? "(no active workbook)"}");
            }
            else
            {
                TryAutoConnect();
            }
        }

        private void PopulateEventsGrid()
        {
            dgvEvents.Rows.Clear();
            foreach (var ev in _events)
                dgvEvents.Rows.Add(ev.Name, ev.Category, ev.Log);
        }

        // ── File selection ───────────────────────────────────────────────────────────
        public void InjectApplication(NetOffice.ExcelApi.Application app)
        {
            if (app == null) return;
            _excelAppNo           = app;
            _applicationInjected  = true;
            try { _excelWbNo = _excelAppNo.ActiveWorkbook; } catch { }
            WireNetOfficeEvents();
            _isConnected = true;
            if (IsHandleCreated)
            {
                btnConnectDisconnect.Text    = "Disconnect";
                btnConnectDisconnect.Enabled = false;
                AppendLog($"Connected (via addin): {_excelWbNo?.Name ?? "(no active workbook)"}");
            }
        }

        private void TryAutoConnect()
        {
            try
            {
                var raw     = WatcherComHelper.GetActiveCom("Excel.Application");
                _excelAppNo = new NetOffice.ExcelApi.Application(null, raw);
                _excelWbNo  = _excelAppNo.ActiveWorkbook;
                WireNetOfficeEvents();
                _isConnected              = true;
                btnConnectDisconnect.Text = "Disconnect";
                AppendLog($"Connected: {_excelWbNo.Name}");
                RefreshOpenDocumentsList();
            }
            catch (Exception ex)
            {
                AppendLog($"Auto-connect failed: {ex.Message}");
            }
        }

        private void btnSelectOutputFile_Click(object sender, EventArgs e)
        {
            using var dlg = new SaveFileDialog
            {
                Title      = "Save JSON Output",
                Filter     = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                DefaultExt = "json",
                FileName   = SaveAsJson_Helper.BuildDefaultPath(_excelWbNo?.FullName, "ETW_MsExcel", cbxJsonWriteMethod.SelectedItem?.ToString())
            };
            if (dlg.ShowDialog(this) == DialogResult.OK)
                txtOutputFilePath.Text = dlg.FileName;
        }

        // ── Connect / Disconnect ─────────────────────────────────────────────────────
        private void btnConnectDisconnect_Click(object sender, EventArgs e)
        {
            if (btnConnectDisconnect.Text == "Connect")
            {
                TryAutoConnect();
            }
            else
            {
                DisconnectAll(quit: false);
                btnConnectDisconnect.Text = "Connect";
                AppendLog("Disconnected.");
            }
        }

        // ── Disconnect / cleanup ─────────────────────────────────────────────────────
        private void DisconnectAll(bool quit)
        {
            try
            {
                UnwireNetOfficeEvents();
                if (quit) { try { _excelWbNo?.Close(); } catch { } try { _excelAppNo?.Quit(); } catch { } }
                _excelWbNo?.Dispose(); _excelWbNo = null;
                if (!_applicationInjected) { _excelAppNo?.Dispose(); }
                _excelAppNo = null;
            }
            catch { }
            _isConnected = false;
            RefreshOpenDocumentsList();
        }

        // ── Events DataGridView ──────────────────────────────────────────────────────
        private void dgvEvents_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvEvents.IsCurrentCellDirty && dgvEvents.CurrentCell is DataGridViewCheckBoxCell)
                dgvEvents.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgvEvents_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != colLog.Index) return;
            _events[e.RowIndex].Log = (bool)(dgvEvents.Rows[e.RowIndex].Cells[colLog.Index].Value ?? false);
        }

        // ── Log helpers ──────────────────────────────────────────────────────────────
        private void btnClearLog_Click(object sender, EventArgs e) => rtbLog.Clear();

        private void LogEvent(string name, string detail)
        {
            var ev = _events.FirstOrDefault(x => x.Name == name);
            if (ev?.Log == true)
                AppendLog($"[Event] {name}{(detail.Length > 0 ? ": " + detail : "")}");
        }

        private void AppendLog(string message)
        {
            // NetOffice COM events can fire while the form is closing; guard against a disposed
            // RichTextBox so AppendText does not resurrect the handle (ObjectDisposedException).
            if (IsDisposed || Disposing) return;
            var box = rtbLog;
            if (box == null || box.IsDisposed) return;
            if (box.InvokeRequired)
            {
                try { box.BeginInvoke(new Action(() => AppendLog(message))); } catch { }
                return;
            }
            try
            {
                box.AppendText($"[{DateTime.Now:HH:mm:ss.fff}] {message}{Environment.NewLine}");
                box.ScrollToCaret();
            }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        // ── Default output path (follows the active workbook) ────────────────────────
        private string? TrySafeWbFullName()
        {
            try { return _excelWbNo?.FullName; } catch { return null; }
        }

        // Recompute the default path when the write method changes so the file name
        // reflects the method (e.g. filename-NetOffice.xlsx.json).
        private void cbxJsonWriteMethod_SelectedIndexChanged(object sender, EventArgs e)
            => UpdateDefaultOutputPath();

        // Refreshes the default .json output path to follow the active workbook and the
        // selected write method (<active file name>-<method>.json). A user-entered custom
        // path is left untouched.
        private void UpdateDefaultOutputPath()
        {
            if (txtOutputFilePath.InvokeRequired)
            {
                try { txtOutputFilePath.BeginInvoke(new Action(UpdateDefaultOutputPath)); } catch { }
                return;
            }
            var newDefault = SaveAsJson_Helper.BuildDefaultPath(TrySafeWbFullName(), "ETW_MsExcel", cbxJsonWriteMethod.SelectedItem?.ToString());
            var current    = txtOutputFilePath.Text.Trim();
            if (current.Length == 0 || current == _lastDefaultOutputPath)
                txtOutputFilePath.Text = newDefault;
            _lastDefaultOutputPath = newDefault;
        }

        // ── Save As JSON ─────────────────────────────────────────────────────────────
        private void btnSaveAsJson_Click(object sender, EventArgs e)
        {
            if (!_isConnected) { AppendLog("ERROR: No workbook connected."); return; }

            var path = txtOutputFilePath.Text.Trim();
            if (string.IsNullOrEmpty(path))
            {
                btnSelectOutputFile_Click(sender, e);
                path = txtOutputFilePath.Text.Trim();
                if (string.IsNullOrEmpty(path)) return;
            }

            var method = cbxJsonWriteMethod.SelectedItem is string s
                && Enum.TryParse<JsonWriteMethod>(s, out var m) ? m : JsonWriteMethod.NetOffice;

            AppendLog($"Save As JSON started ({method}) → {path}");
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                MsOfficeSaveAsJsonWriterProvider.Require().WriteExcelJson(
                    method.ToString(), _excelWbNo, _excelAppNo, path, AppendLog);
                stopwatch.Stop();
                AppendLog($"✔ Save As JSON complete ({method}) in {Duration_Helper.Format(stopwatch.Elapsed)} → {path}");
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                AppendLog($"✘ Save As JSON failed after {Duration_Helper.Format(stopwatch.Elapsed)}: {ex.Message}");
            }
        }

        // ── Form closing ─────────────────────────────────────────────────────────────
        private void MsExcel_Watcher_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            DisconnectAll(quit: false);
        }

        private void RefreshOpenDocumentsList()
        {
            if (_cbxDocuments == null) return;
            if (_cbxDocuments.InvokeRequired)
            {
                _cbxDocuments.Invoke(new Action(RefreshOpenDocumentsList));
                return;
            }

            try
            {
                var selected = _excelWbNo?.Name ?? _cbxDocuments.SelectedItem as string;
                _cbxDocuments.Items.Clear();
                if (_excelAppNo != null)
                {
                    foreach (NetOffice.ExcelApi.Workbook wb in _excelAppNo.Workbooks)
                    {
                        try { _cbxDocuments.Items.Add(wb.Name); } catch { }
                    }
                }

                _cbxDocuments.Enabled = _cbxDocuments.Items.Count > 0;
                if (_cbxDocuments.Items.Count > 0)
                {
                    if (selected != null && _cbxDocuments.Items.Contains(selected))
                        _cbxDocuments.SelectedItem = selected;
                    else
                        _cbxDocuments.SelectedIndex = 0;
                }
                else
                {
                    _cbxDocuments.SelectedIndex = -1;
                }
            }
            catch { }
        }

        private void cbxDocuments_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_cbxDocuments?.SelectedItem is not string wbName || _excelAppNo == null) return;
            try
            {
                foreach (NetOffice.ExcelApi.Workbook wb in _excelAppNo.Workbooks)
                {
                    if (wb.Name == wbName)
                    {
                        _excelWbNo = wb;
                        wb.Activate();
                        UpdateDefaultOutputPath();
                        AppendLog($"Selected workbook: {wbName}");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Error selecting workbook: {ex.Message}");
            }
        }

        // ── NetOffice event handlers ─────────────────────────────────────────────────
        // Calculation
        private void OnNo_AfterCalculate()                                                                                       => LogEvent("AfterCalculate", "");
        // Chart
        private void OnNo_WorkbookNewChart(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Chart ch)                          => LogEvent("WorkbookNewChart", wb.Name);
        // Data
        private void OnNo_WorkbookRowsetComplete(NetOffice.ExcelApi.Workbook wb, string description, string sheet, bool success) => LogEvent("WorkbookRowsetComplete", wb.Name);
        // Data Model
        private void OnNo_WorkbookModelChange(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.ModelChanges changes)           => LogEvent("WorkbookModelChange", wb.Name);
        // PivotTable
        private void OnNo_SheetPivotTableAfterValueChange(NetOffice.ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange) => LogEvent("SheetPivotTableAfterValueChange", "");
        private void OnNo_SheetPivotTableBeforeAllocateChanges(NetOffice.ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, int valueChangeStart, int valueChangeEnd, ref bool cancel) => LogEvent("SheetPivotTableBeforeAllocateChanges", "");
        private void OnNo_SheetPivotTableBeforeCommitChanges(NetOffice.ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, int valueChangeStart, int valueChangeEnd, ref bool cancel) => LogEvent("SheetPivotTableBeforeCommitChanges", "");
        private void OnNo_SheetPivotTableBeforeDiscardChanges(NetOffice.ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, int valueChangeStart, int valueChangeEnd) => LogEvent("SheetPivotTableBeforeDiscardChanges", "");
        private void OnNo_SheetPivotTableUpdate(NetOffice.ICOMObject sh, NetOffice.ExcelApi.PivotTable target)                   => LogEvent("SheetPivotTableUpdate", "");
        private void OnNo_WorkbookPivotTableCloseConnection(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target) => LogEvent("WorkbookPivotTableCloseConnection", wb.Name);
        private void OnNo_WorkbookPivotTableOpenConnection(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target)  => LogEvent("WorkbookPivotTableOpenConnection", wb.Name);
        // Table
        private void OnNo_SheetTableUpdate(NetOffice.ICOMObject sh, NetOffice.ExcelApi.TableObject target)                       => LogEvent("SheetTableUpdate", "");
        // Workbook
        private void OnNo_NewWorkbook(NetOffice.ExcelApi.Workbook wb)
        {
            RefreshOpenDocumentsList();
            LogEvent("NewWorkbook", wb.Name);
        }
        private void OnNo_WorkbookActivate(NetOffice.ExcelApi.Workbook wb)
        {
            _excelWbNo = wb;
            RefreshOpenDocumentsList();
            UpdateDefaultOutputPath();
            LogEvent("WorkbookActivate", wb.Name);
        }
        private void OnNo_WorkbookAfterSave(NetOffice.ExcelApi.Workbook wb, bool success)                                        => LogEvent("WorkbookAfterSave", wb.Name);
        private void OnNo_WorkbookBeforeClose(NetOffice.ExcelApi.Workbook wb, ref bool cancel)                                   => LogEvent("WorkbookBeforeClose", wb.Name);
        private void OnNo_WorkbookBeforePrint(NetOffice.ExcelApi.Workbook wb, ref bool cancel)                                   => LogEvent("WorkbookBeforePrint", wb.Name);
        private void OnNo_WorkbookBeforeSave(NetOffice.ExcelApi.Workbook wb, bool saveAsUi, ref bool cancel)                     => LogEvent("WorkbookBeforeSave", wb.Name);
        private void OnNo_WorkbookDeactivate(NetOffice.ExcelApi.Workbook wb)                                                     => LogEvent("WorkbookDeactivate", wb.Name);
        private void OnNo_WorkbookOpen(NetOffice.ExcelApi.Workbook wb)
        {
            RefreshOpenDocumentsList();
            LogEvent("WorkbookOpen", wb.Name);
        }
        private void OnNo_WorkbookSync(NetOffice.ExcelApi.Workbook wb, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType)  => LogEvent("WorkbookSync", wb.Name);
        // Workbook/Add-in
        private void OnNo_WorkbookAddinInstall(NetOffice.ExcelApi.Workbook wb)                                                   => LogEvent("WorkbookAddinInstall", wb.Name);
        private void OnNo_WorkbookAddinUninstall(NetOffice.ExcelApi.Workbook wb)                                                 => LogEvent("WorkbookAddinUninstall", wb.Name);
        // Worksheet
        private void OnNo_SheetActivate(NetOffice.ICOMObject sh)                                                                 => LogEvent("SheetActivate", "");
        private void OnNo_SheetBeforeDelete(NetOffice.ICOMObject sh)                                                             => LogEvent("SheetBeforeDelete", "");
        private void OnNo_SheetBeforeDoubleClick(NetOffice.ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel)      => LogEvent("SheetBeforeDoubleClick", "");
        private void OnNo_SheetBeforeRightClick(NetOffice.ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel)       => LogEvent("SheetBeforeRightClick", "");
        private void OnNo_SheetCalculate(NetOffice.ICOMObject sh)                                                                => LogEvent("SheetCalculate", "");
        private void OnNo_SheetChange(NetOffice.ICOMObject sh, NetOffice.ExcelApi.Range target)                                  => LogEvent("SheetChange", target.Address);
        private void OnNo_SheetDeactivate(NetOffice.ICOMObject sh)                                                               => LogEvent("SheetDeactivate", "");
        private void OnNo_SheetFollowHyperlink(NetOffice.ICOMObject sh, NetOffice.ExcelApi.Hyperlink target)                     => LogEvent("SheetFollowHyperlink", "");
        private void OnNo_SheetLensGalleryRenderComplete(NetOffice.ICOMObject sh)                                                => LogEvent("SheetLensGalleryRenderComplete", "");
        private void OnNo_SheetSelectionChange(NetOffice.ICOMObject sh, NetOffice.ExcelApi.Range target)                         => LogEvent("SheetSelectionChange", "");
        private void OnNo_WorkbookNewSheet(NetOffice.ExcelApi.Workbook wb, NetOffice.ICOMObject sh)                              => LogEvent("WorkbookNewSheet", wb.Name);
        // XML
        private void OnNo_WorkbookAfterXmlExport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result) => LogEvent("WorkbookAfterXmlExport", wb.Name);
        private void OnNo_WorkbookAfterXmlImport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result) => LogEvent("WorkbookAfterXmlImport", wb.Name);
        private void OnNo_WorkbookBeforeXmlExport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, ref bool cancel) => LogEvent("WorkbookBeforeXmlExport", wb.Name);
        private void OnNo_WorkbookBeforeXmlImport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, ref bool cancel) => LogEvent("WorkbookBeforeXmlImport", wb.Name);
        // Window
        private void OnNo_WindowActivate(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)
        {
            _excelWbNo = wb;
            RefreshOpenDocumentsList();
            UpdateDefaultOutputPath();
            LogEvent("WindowActivate", wb.Name);
        }
        private void OnNo_WindowDeactivate(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)                        => LogEvent("WindowDeactivate", wb.Name);
        private void OnNo_WindowResize(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)                            => LogEvent("WindowResize", wb.Name);
        // Protected View
        private void OnNo_ProtectedViewWindowOpen(NetOffice.ExcelApi.ProtectedViewWindow pvw)                                    => LogEvent("ProtectedViewWindowOpen", "");
        private void OnNo_ProtectedViewWindowBeforeEdit(NetOffice.ExcelApi.ProtectedViewWindow pvw, ref bool cancel)             => LogEvent("ProtectedViewWindowBeforeEdit", "");
        private void OnNo_ProtectedViewWindowBeforeClose(NetOffice.ExcelApi.ProtectedViewWindow pvw, NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason reason, ref bool cancel) => LogEvent("ProtectedViewWindowBeforeClose", "");
        private void OnNo_ProtectedViewWindowResize(NetOffice.ExcelApi.ProtectedViewWindow pvw)                                  => LogEvent("ProtectedViewWindowResize", "");
        private void OnNo_ProtectedViewWindowActivate(NetOffice.ExcelApi.ProtectedViewWindow pvw)                                => LogEvent("ProtectedViewWindowActivate", "");
        private void OnNo_ProtectedViewWindowDeactivate(NetOffice.ExcelApi.ProtectedViewWindow pvw)                              => LogEvent("ProtectedViewWindowDeactivate", "");

        private void WireNetOfficeEvents()
        {
            if (_excelAppNo == null) return;
            // Calculation
            _excelAppNo.AfterCalculateEvent                        += OnNo_AfterCalculate;
            // Chart
            _excelAppNo.WorkbookNewChartEvent                      += OnNo_WorkbookNewChart;
            // Data
            _excelAppNo.WorkbookRowsetCompleteEvent                += OnNo_WorkbookRowsetComplete;
            // Data Model
            _excelAppNo.WorkbookModelChangeEvent                   += OnNo_WorkbookModelChange;
            // PivotTable
            _excelAppNo.SheetPivotTableAfterValueChangeEvent       += OnNo_SheetPivotTableAfterValueChange;
            _excelAppNo.SheetPivotTableBeforeAllocateChangesEvent  += OnNo_SheetPivotTableBeforeAllocateChanges;
            _excelAppNo.SheetPivotTableBeforeCommitChangesEvent    += OnNo_SheetPivotTableBeforeCommitChanges;
            _excelAppNo.SheetPivotTableBeforeDiscardChangesEvent   += OnNo_SheetPivotTableBeforeDiscardChanges;
            _excelAppNo.SheetPivotTableUpdateEvent                 += OnNo_SheetPivotTableUpdate;
            _excelAppNo.WorkbookPivotTableCloseConnectionEvent     += OnNo_WorkbookPivotTableCloseConnection;
            _excelAppNo.WorkbookPivotTableOpenConnectionEvent      += OnNo_WorkbookPivotTableOpenConnection;
            // Table
            _excelAppNo.SheetTableUpdateEvent                      += OnNo_SheetTableUpdate;
            // Workbook
            _excelAppNo.NewWorkbookEvent                           += OnNo_NewWorkbook;
            _excelAppNo.WorkbookActivateEvent                      += OnNo_WorkbookActivate;
            _excelAppNo.WorkbookAfterSaveEvent                     += OnNo_WorkbookAfterSave;
            _excelAppNo.WorkbookBeforeCloseEvent                   += OnNo_WorkbookBeforeClose;
            _excelAppNo.WorkbookBeforePrintEvent                   += OnNo_WorkbookBeforePrint;
            _excelAppNo.WorkbookBeforeSaveEvent                    += OnNo_WorkbookBeforeSave;
            _excelAppNo.WorkbookDeactivateEvent                    += OnNo_WorkbookDeactivate;
            _excelAppNo.WorkbookOpenEvent                          += OnNo_WorkbookOpen;
            _excelAppNo.WorkbookSyncEvent                          += OnNo_WorkbookSync;
            // Workbook/Add-in
            _excelAppNo.WorkbookAddinInstallEvent                  += OnNo_WorkbookAddinInstall;
            _excelAppNo.WorkbookAddinUninstallEvent                += OnNo_WorkbookAddinUninstall;
            // Worksheet
            _excelAppNo.SheetActivateEvent                         += OnNo_SheetActivate;
            _excelAppNo.SheetBeforeDeleteEvent                     += OnNo_SheetBeforeDelete;
            _excelAppNo.SheetBeforeDoubleClickEvent                += OnNo_SheetBeforeDoubleClick;
            _excelAppNo.SheetBeforeRightClickEvent                 += OnNo_SheetBeforeRightClick;
            _excelAppNo.SheetCalculateEvent                        += OnNo_SheetCalculate;
            _excelAppNo.SheetChangeEvent                           += OnNo_SheetChange;
            _excelAppNo.SheetDeactivateEvent                       += OnNo_SheetDeactivate;
            _excelAppNo.SheetFollowHyperlinkEvent                  += OnNo_SheetFollowHyperlink;
            _excelAppNo.SheetLensGalleryRenderCompleteEvent        += OnNo_SheetLensGalleryRenderComplete;
            _excelAppNo.SheetSelectionChangeEvent                  += OnNo_SheetSelectionChange;
            _excelAppNo.WorkbookNewSheetEvent                      += OnNo_WorkbookNewSheet;
            // XML
            _excelAppNo.WorkbookAfterXmlExportEvent                += OnNo_WorkbookAfterXmlExport;
            _excelAppNo.WorkbookAfterXmlImportEvent                += OnNo_WorkbookAfterXmlImport;
            _excelAppNo.WorkbookBeforeXmlExportEvent               += OnNo_WorkbookBeforeXmlExport;
            _excelAppNo.WorkbookBeforeXmlImportEvent               += OnNo_WorkbookBeforeXmlImport;
            // Window
            _excelAppNo.WindowActivateEvent                        += OnNo_WindowActivate;
            _excelAppNo.WindowDeactivateEvent                      += OnNo_WindowDeactivate;
            _excelAppNo.WindowResizeEvent                          += OnNo_WindowResize;
            // Protected View
            _excelAppNo.ProtectedViewWindowOpenEvent               += OnNo_ProtectedViewWindowOpen;
            _excelAppNo.ProtectedViewWindowBeforeEditEvent         += OnNo_ProtectedViewWindowBeforeEdit;
            _excelAppNo.ProtectedViewWindowBeforeCloseEvent        += OnNo_ProtectedViewWindowBeforeClose;
            _excelAppNo.ProtectedViewWindowResizeEvent             += OnNo_ProtectedViewWindowResize;
            _excelAppNo.ProtectedViewWindowActivateEvent           += OnNo_ProtectedViewWindowActivate;
            _excelAppNo.ProtectedViewWindowDeactivateEvent         += OnNo_ProtectedViewWindowDeactivate;
        }

        private void UnwireNetOfficeEvents()
        {
            if (_excelAppNo == null) return;
            try
            {
                // Calculation
                _excelAppNo.AfterCalculateEvent                        -= OnNo_AfterCalculate;
                // Chart
                _excelAppNo.WorkbookNewChartEvent                      -= OnNo_WorkbookNewChart;
                // Data
                _excelAppNo.WorkbookRowsetCompleteEvent                -= OnNo_WorkbookRowsetComplete;
                // Data Model
                _excelAppNo.WorkbookModelChangeEvent                   -= OnNo_WorkbookModelChange;
                // PivotTable
                _excelAppNo.SheetPivotTableAfterValueChangeEvent       -= OnNo_SheetPivotTableAfterValueChange;
                _excelAppNo.SheetPivotTableBeforeAllocateChangesEvent  -= OnNo_SheetPivotTableBeforeAllocateChanges;
                _excelAppNo.SheetPivotTableBeforeCommitChangesEvent    -= OnNo_SheetPivotTableBeforeCommitChanges;
                _excelAppNo.SheetPivotTableBeforeDiscardChangesEvent   -= OnNo_SheetPivotTableBeforeDiscardChanges;
                _excelAppNo.SheetPivotTableUpdateEvent                 -= OnNo_SheetPivotTableUpdate;
                _excelAppNo.WorkbookPivotTableCloseConnectionEvent     -= OnNo_WorkbookPivotTableCloseConnection;
                _excelAppNo.WorkbookPivotTableOpenConnectionEvent      -= OnNo_WorkbookPivotTableOpenConnection;
                // Table
                _excelAppNo.SheetTableUpdateEvent                      -= OnNo_SheetTableUpdate;
                // Workbook
                _excelAppNo.NewWorkbookEvent                           -= OnNo_NewWorkbook;
                _excelAppNo.WorkbookActivateEvent                      -= OnNo_WorkbookActivate;
                _excelAppNo.WorkbookAfterSaveEvent                     -= OnNo_WorkbookAfterSave;
                _excelAppNo.WorkbookBeforeCloseEvent                   -= OnNo_WorkbookBeforeClose;
                _excelAppNo.WorkbookBeforePrintEvent                   -= OnNo_WorkbookBeforePrint;
                _excelAppNo.WorkbookBeforeSaveEvent                    -= OnNo_WorkbookBeforeSave;
                _excelAppNo.WorkbookDeactivateEvent                    -= OnNo_WorkbookDeactivate;
                _excelAppNo.WorkbookOpenEvent                          -= OnNo_WorkbookOpen;
                _excelAppNo.WorkbookSyncEvent                          -= OnNo_WorkbookSync;
                // Workbook/Add-in
                _excelAppNo.WorkbookAddinInstallEvent                  -= OnNo_WorkbookAddinInstall;
                _excelAppNo.WorkbookAddinUninstallEvent                -= OnNo_WorkbookAddinUninstall;
                // Worksheet
                _excelAppNo.SheetActivateEvent                         -= OnNo_SheetActivate;
                _excelAppNo.SheetBeforeDeleteEvent                     -= OnNo_SheetBeforeDelete;
                _excelAppNo.SheetBeforeDoubleClickEvent                -= OnNo_SheetBeforeDoubleClick;
                _excelAppNo.SheetBeforeRightClickEvent                 -= OnNo_SheetBeforeRightClick;
                _excelAppNo.SheetCalculateEvent                        -= OnNo_SheetCalculate;
                _excelAppNo.SheetChangeEvent                           -= OnNo_SheetChange;
                _excelAppNo.SheetDeactivateEvent                       -= OnNo_SheetDeactivate;
                _excelAppNo.SheetFollowHyperlinkEvent                  -= OnNo_SheetFollowHyperlink;
                _excelAppNo.SheetLensGalleryRenderCompleteEvent        -= OnNo_SheetLensGalleryRenderComplete;
                _excelAppNo.SheetSelectionChangeEvent                  -= OnNo_SheetSelectionChange;
                _excelAppNo.WorkbookNewSheetEvent                      -= OnNo_WorkbookNewSheet;
                // XML
                _excelAppNo.WorkbookAfterXmlExportEvent                -= OnNo_WorkbookAfterXmlExport;
                _excelAppNo.WorkbookAfterXmlImportEvent                -= OnNo_WorkbookAfterXmlImport;
                _excelAppNo.WorkbookBeforeXmlExportEvent               -= OnNo_WorkbookBeforeXmlExport;
                _excelAppNo.WorkbookBeforeXmlImportEvent               -= OnNo_WorkbookBeforeXmlImport;
                // Window
                _excelAppNo.WindowActivateEvent                        -= OnNo_WindowActivate;
                _excelAppNo.WindowDeactivateEvent                      -= OnNo_WindowDeactivate;
                _excelAppNo.WindowResizeEvent                          -= OnNo_WindowResize;
                // Protected View
                _excelAppNo.ProtectedViewWindowOpenEvent               -= OnNo_ProtectedViewWindowOpen;
                _excelAppNo.ProtectedViewWindowBeforeEditEvent         -= OnNo_ProtectedViewWindowBeforeEdit;
                _excelAppNo.ProtectedViewWindowBeforeCloseEvent        -= OnNo_ProtectedViewWindowBeforeClose;
                _excelAppNo.ProtectedViewWindowResizeEvent             -= OnNo_ProtectedViewWindowResize;
                _excelAppNo.ProtectedViewWindowActivateEvent           -= OnNo_ProtectedViewWindowActivate;
                _excelAppNo.ProtectedViewWindowDeactivateEvent         -= OnNo_ProtectedViewWindowDeactivate;
            }
            catch { }
        }
    }
}
