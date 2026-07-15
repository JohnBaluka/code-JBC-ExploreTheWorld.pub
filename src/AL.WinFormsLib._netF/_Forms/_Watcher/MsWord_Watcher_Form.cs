using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using JBC.ExploreTheWorld.CL;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class MsWord_Watcher_Form : Form
    {
        // ── State ────────────────────────────────────────────────────────────────────
        private readonly List<MsOfficeEvent_Record> _events = MsOfficeEvents_Repo.ForWord();
        private ComboBox? _cbxDocuments;
        private bool _isConnected;
        private bool _applicationInjected;

        // Last auto-generated default output path — used so a user-entered custom path is preserved
        // when the active document changes.
        private string _lastDefaultOutputPath = string.Empty;

        // NetOffice
        private NetOffice.WordApi.Application? _wordAppNo;
        private NetOffice.WordApi.Document?    _wordDocNo;

        // ── Constructor ──────────────────────────────────────────────────────────────
        public MsWord_Watcher_Form()
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
        private void MsWord_Watcher_Form_Load(object sender, EventArgs e)
        {
            _lastDefaultOutputPath = SaveAsJson_Helper.BuildDefaultPath(TrySafeDocFullName(), "ETW_MsWord", cbxJsonWriteMethod.SelectedItem?.ToString());
            txtOutputFilePath.Text = _lastDefaultOutputPath;
            cbxJsonWriteMethod.Items.AddRange(Enum.GetNames(typeof(JsonWriteMethod)));
            cbxJsonWriteMethod.SelectedItem = JsonWriteMethod.NetOffice.ToString();
            PopulateEventsGrid();
            if (_applicationInjected)
            {
                btnConnectDisconnect.Text    = "Disconnect";
                btnConnectDisconnect.Enabled = false;
                AppendLog($"Connected (via addin): {_wordDocNo?.Name ?? "(no active document)"}");
                RefreshOpenDocumentsList();
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

        // ── Inject (from addin) / auto-connect ───────────────────────────────────────
        public void InjectApplication(NetOffice.WordApi.Application app)
        {
            if (app == null) return;
            _wordAppNo            = app;
            _applicationInjected  = true;
            try { _wordDocNo = _wordAppNo.ActiveDocument; } catch { }
            WireNetOfficeEvents();
            _isConnected = true;
            if (IsHandleCreated)
            {
                btnConnectDisconnect.Text    = "Disconnect";
                btnConnectDisconnect.Enabled = false;
                AppendLog($"Connected (via addin): {_wordDocNo?.Name ?? "(no active document)"}");
                RefreshOpenDocumentsList();
            }
        }

        private void TryAutoConnect()
        {
            try
            {
                var raw    = WatcherComHelper.GetActiveCom("Word.Application");
                _wordAppNo = new NetOffice.WordApi.Application(null, raw);
                _wordDocNo = _wordAppNo.ActiveDocument;
                WireNetOfficeEvents();
                _isConnected              = true;
                btnConnectDisconnect.Text = "Disconnect";
                AppendLog($"Connected: {_wordDocNo.Name}");
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
                FileName   = SaveAsJson_Helper.BuildDefaultPath(_wordDocNo?.FullName, "ETW_MsWord", cbxJsonWriteMethod.SelectedItem?.ToString())
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
                if (quit) { try { _wordDocNo?.Close(); } catch { } try { _wordAppNo?.Quit(); } catch { } }
                _wordDocNo?.Dispose(); _wordDocNo = null;
                if (!_applicationInjected) { _wordAppNo?.Dispose(); }
                _wordAppNo = null;
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

        // ── Default output path (follows the active document) ────────────────────────
        private string? TrySafeDocFullName()
        {
            try { return _wordDocNo?.FullName; } catch { return null; }
        }

        // Recompute the default path when the write method changes so the file name
        // reflects the method (e.g. filename-NetOffice.docx.json).
        private void cbxJsonWriteMethod_SelectedIndexChanged(object sender, EventArgs e)
            => UpdateDefaultOutputPath();

        // Refreshes the default .json output path to follow the active document and the
        // selected write method (<active file name>-<method>.json). A user-entered custom
        // path is left untouched.
        private void UpdateDefaultOutputPath()
        {
            if (txtOutputFilePath.InvokeRequired)
            {
                try { txtOutputFilePath.BeginInvoke(new Action(UpdateDefaultOutputPath)); } catch { }
                return;
            }
            var newDefault = SaveAsJson_Helper.BuildDefaultPath(TrySafeDocFullName(), "ETW_MsWord", cbxJsonWriteMethod.SelectedItem?.ToString());
            var current    = txtOutputFilePath.Text.Trim();
            if (current.Length == 0 || current == _lastDefaultOutputPath)
                txtOutputFilePath.Text = newDefault;
            _lastDefaultOutputPath = newDefault;
        }

        // ── Save As JSON ─────────────────────────────────────────────────────────────
        private void btnSaveAsJson_Click(object sender, EventArgs e)
        {
            if (!_isConnected) { AppendLog("ERROR: No document connected."); return; }

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
                MsOfficeSaveAsJsonWriterProvider.Require().WriteWordJson(
                    method.ToString(), _wordDocNo, _wordAppNo, path, AppendLog);
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
        private void MsWord_Watcher_Form_FormClosing(object sender, FormClosingEventArgs e)
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
                var selected = _wordDocNo?.Name ?? _cbxDocuments.SelectedItem as string;
                _cbxDocuments.Items.Clear();
                if (_wordAppNo != null)
                {
                    foreach (NetOffice.WordApi.Document doc in _wordAppNo.Documents)
                    {
                        try { _cbxDocuments.Items.Add(doc.Name); } catch { }
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
            if (_cbxDocuments?.SelectedItem is not string docName || _wordAppNo == null) return;
            try
            {
                foreach (NetOffice.WordApi.Document doc in _wordAppNo.Documents)
                {
                    if (doc.Name == docName)
                    {
                        _wordDocNo = doc;
                        doc.Activate();
                        UpdateDefaultOutputPath();
                        AppendLog($"Selected document: {docName}");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Error selecting document: {ex.Message}");
            }
        }

        // ── DotNetInterop PIA event handlers ─────────────────────────────────────────
        // (removed — DotNetInterop option eliminated; use NetOffice)

        // ── NetOffice event handlers ─────────────────────────────────────────────────
        // Application
        private void OnNo_Quit()                                                                        => LogEvent("Quit", "");
        // Document
        private void OnNo_NewDocument(NetOffice.WordApi.Document doc)
        {
            RefreshOpenDocumentsList();
            LogEvent("NewDocument", doc.Name);
        }
        private void OnNo_DocumentOpen(NetOffice.WordApi.Document doc)
        {
            RefreshOpenDocumentsList();
            LogEvent("DocumentOpen", doc.Name);
        }
        private void OnNo_DocumentBeforeClose(NetOffice.WordApi.Document doc, ref bool cancel)
        {
            RefreshOpenDocumentsList();
            LogEvent("DocumentBeforeClose", doc.Name);
        }
        private void OnNo_DocumentBeforePrint(NetOffice.WordApi.Document doc, ref bool cancel)          => LogEvent("DocumentBeforePrint", doc.Name);
        private void OnNo_DocumentBeforeSave(NetOffice.WordApi.Document doc, ref bool saveAsUI, ref bool cancel) => LogEvent("DocumentBeforeSave", doc.Name);
        private void OnNo_DocumentChange()                                                              => LogEvent("DocumentChange", "");
        private void OnNo_DocumentSync(NetOffice.WordApi.Document doc, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType) => LogEvent("DocumentSync", doc.Name);
        // E-Postage
        private void OnNo_EPostageInsert(NetOffice.WordApi.Document doc)                                => LogEvent("EPostageInsert", doc.Name);
        private void OnNo_EPostageInsertEx(NetOffice.WordApi.Document doc, int cpDeliveryAddrStart, int cpDeliveryAddrEnd, int cpReturnAddrStart, int cpReturnAddrEnd, int xaWidth, int yaHeight, string bstrPrinterName, string bstrPaperFeed, bool fPrint, ref bool fCancel) => LogEvent("EPostageInsertEx", doc.Name);
        private void OnNo_EPostagePropertyDialog(NetOffice.WordApi.Document doc)                        => LogEvent("EPostagePropertyDialog", doc.Name);
        // Mail Merge
        private void OnNo_MailMergeAfterMerge(NetOffice.WordApi.Document doc, NetOffice.WordApi.Document docResult) => LogEvent("MailMergeAfterMerge", doc.Name);
        private void OnNo_MailMergeAfterRecordMerge(NetOffice.WordApi.Document doc)                     => LogEvent("MailMergeAfterRecordMerge", doc.Name);
        private void OnNo_MailMergeBeforeMerge(NetOffice.WordApi.Document doc, int startRecord, int endRecord, ref bool cancel) => LogEvent("MailMergeBeforeMerge", doc.Name);
        private void OnNo_MailMergeBeforeRecordMerge(NetOffice.WordApi.Document doc, ref bool cancel)   => LogEvent("MailMergeBeforeRecordMerge", doc.Name);
        private void OnNo_MailMergeDataSourceLoad(NetOffice.WordApi.Document doc)                       => LogEvent("MailMergeDataSourceLoad", doc.Name);
        private void OnNo_MailMergeDataSourceValidate(NetOffice.WordApi.Document doc, ref bool handled) => LogEvent("MailMergeDataSourceValidate", doc.Name);
        private void OnNo_MailMergeDataSourceValidate2(NetOffice.WordApi.Document doc, ref bool handled) => LogEvent("MailMergeDataSourceValidate2", doc.Name);
        private void OnNo_MailMergeWizardSendToCustom(NetOffice.WordApi.Document doc)                   => LogEvent("MailMergeWizardSendToCustom", doc.Name);
        private void OnNo_MailMergeWizardStateChange(NetOffice.WordApi.Document doc, ref int fromState, ref int toState, ref bool handled) => LogEvent("MailMergeWizardStateChange", doc.Name);
        // XML
        private void OnNo_XMLSelectionChange(NetOffice.WordApi.Selection sel, NetOffice.WordApi.XMLNode oldXMLNode, NetOffice.WordApi.XMLNode newXMLNode, ref int reason) => LogEvent("XMLSelectionChange", "");
        private void OnNo_XMLValidationError(NetOffice.WordApi.XMLNode xmlNode)                         => LogEvent("XMLValidationError", "");
        // Window
        private void OnNo_WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
        {
            _wordDocNo = doc;
            RefreshOpenDocumentsList();
            UpdateDefaultOutputPath();
            LogEvent("WindowActivate", doc.Name);
        }
        private void OnNo_WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn) => LogEvent("WindowDeactivate", doc.Name);
        private void OnNo_WindowSelectionChange(NetOffice.WordApi.Selection sel)                        => LogEvent("WindowSelectionChange", "");
        private void OnNo_WindowBeforeRightClick(NetOffice.WordApi.Selection sel, ref bool cancel)      => LogEvent("WindowBeforeRightClick", "");
        private void OnNo_WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, ref bool cancel)     => LogEvent("WindowBeforeDoubleClick", "");
        private void OnNo_WindowSize(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)       => LogEvent("WindowSize", doc.Name);
        // Protected View
        private void OnNo_ProtectedViewWindowOpen(NetOffice.WordApi.ProtectedViewWindow pvWindow)                                 => LogEvent("ProtectedViewWindowOpen", "");
        private void OnNo_ProtectedViewWindowBeforeEdit(NetOffice.WordApi.ProtectedViewWindow pvWindow, ref bool cancel)          => LogEvent("ProtectedViewWindowBeforeEdit", "");
        private void OnNo_ProtectedViewWindowBeforeClose(NetOffice.WordApi.ProtectedViewWindow pvWindow, int closeReason, ref bool cancel) => LogEvent("ProtectedViewWindowBeforeClose", "");
        private void OnNo_ProtectedViewWindowSize(NetOffice.WordApi.ProtectedViewWindow pvWindow)                                 => LogEvent("ProtectedViewWindowSize", "");
        private void OnNo_ProtectedViewWindowActivate(NetOffice.WordApi.ProtectedViewWindow pvWindow)                             => LogEvent("ProtectedViewWindowActivate", "");
        private void OnNo_ProtectedViewWindowDeactivate(NetOffice.WordApi.ProtectedViewWindow pvWindow)                           => LogEvent("ProtectedViewWindowDeactivate", "");

        private void WireNetOfficeEvents()
        {
            if (_wordAppNo == null) return;
            // Application
            _wordAppNo.QuitEvent                               += OnNo_Quit;
            // Document
            _wordAppNo.NewDocumentEvent                        += OnNo_NewDocument;
            _wordAppNo.DocumentOpenEvent                       += OnNo_DocumentOpen;
            _wordAppNo.DocumentBeforeCloseEvent                += OnNo_DocumentBeforeClose;
            _wordAppNo.DocumentBeforePrintEvent                += OnNo_DocumentBeforePrint;
            _wordAppNo.DocumentBeforeSaveEvent                 += OnNo_DocumentBeforeSave;
            _wordAppNo.DocumentChangeEvent                     += OnNo_DocumentChange;
            _wordAppNo.DocumentSyncEvent                       += OnNo_DocumentSync;
            // E-Postage
            _wordAppNo.EPostageInsertEvent                     += OnNo_EPostageInsert;
            _wordAppNo.EPostageInsertExEvent                   += OnNo_EPostageInsertEx;
            _wordAppNo.EPostagePropertyDialogEvent             += OnNo_EPostagePropertyDialog;
            // Mail Merge
            _wordAppNo.MailMergeAfterMergeEvent                += OnNo_MailMergeAfterMerge;
            _wordAppNo.MailMergeAfterRecordMergeEvent          += OnNo_MailMergeAfterRecordMerge;
            _wordAppNo.MailMergeBeforeMergeEvent               += OnNo_MailMergeBeforeMerge;
            _wordAppNo.MailMergeBeforeRecordMergeEvent         += OnNo_MailMergeBeforeRecordMerge;
            _wordAppNo.MailMergeDataSourceLoadEvent            += OnNo_MailMergeDataSourceLoad;
            _wordAppNo.MailMergeDataSourceValidateEvent        += OnNo_MailMergeDataSourceValidate;
            _wordAppNo.MailMergeDataSourceValidate2Event       += OnNo_MailMergeDataSourceValidate2;
            _wordAppNo.MailMergeWizardSendToCustomEvent        += OnNo_MailMergeWizardSendToCustom;
            _wordAppNo.MailMergeWizardStateChangeEvent         += OnNo_MailMergeWizardStateChange;
            // XML
            _wordAppNo.XMLSelectionChangeEvent                 += OnNo_XMLSelectionChange;
            _wordAppNo.XMLValidationErrorEvent                 += OnNo_XMLValidationError;
            // Window
            _wordAppNo.WindowActivateEvent                     += OnNo_WindowActivate;
            _wordAppNo.WindowDeactivateEvent                   += OnNo_WindowDeactivate;
            _wordAppNo.WindowSelectionChangeEvent              += OnNo_WindowSelectionChange;
            _wordAppNo.WindowBeforeRightClickEvent             += OnNo_WindowBeforeRightClick;
            _wordAppNo.WindowBeforeDoubleClickEvent            += OnNo_WindowBeforeDoubleClick;
            _wordAppNo.WindowSizeEvent                         += OnNo_WindowSize;
            // Protected View
            _wordAppNo.ProtectedViewWindowOpenEvent            += OnNo_ProtectedViewWindowOpen;
            _wordAppNo.ProtectedViewWindowBeforeEditEvent      += OnNo_ProtectedViewWindowBeforeEdit;
            _wordAppNo.ProtectedViewWindowBeforeCloseEvent     += OnNo_ProtectedViewWindowBeforeClose;
            _wordAppNo.ProtectedViewWindowSizeEvent            += OnNo_ProtectedViewWindowSize;
            _wordAppNo.ProtectedViewWindowActivateEvent        += OnNo_ProtectedViewWindowActivate;
            _wordAppNo.ProtectedViewWindowDeactivateEvent      += OnNo_ProtectedViewWindowDeactivate;
        }

        private void UnwireNetOfficeEvents()
        {
            if (_wordAppNo == null) return;
            try
            {
                // Application
                _wordAppNo.QuitEvent                               -= OnNo_Quit;
                // Document
                _wordAppNo.NewDocumentEvent                        -= OnNo_NewDocument;
                _wordAppNo.DocumentOpenEvent                       -= OnNo_DocumentOpen;
                _wordAppNo.DocumentBeforeCloseEvent                -= OnNo_DocumentBeforeClose;
                _wordAppNo.DocumentBeforePrintEvent                -= OnNo_DocumentBeforePrint;
                _wordAppNo.DocumentBeforeSaveEvent                 -= OnNo_DocumentBeforeSave;
                _wordAppNo.DocumentChangeEvent                     -= OnNo_DocumentChange;
                _wordAppNo.DocumentSyncEvent                       -= OnNo_DocumentSync;
                // E-Postage
                _wordAppNo.EPostageInsertEvent                     -= OnNo_EPostageInsert;
                _wordAppNo.EPostageInsertExEvent                   -= OnNo_EPostageInsertEx;
                _wordAppNo.EPostagePropertyDialogEvent             -= OnNo_EPostagePropertyDialog;
                // Mail Merge
                _wordAppNo.MailMergeAfterMergeEvent                -= OnNo_MailMergeAfterMerge;
                _wordAppNo.MailMergeAfterRecordMergeEvent          -= OnNo_MailMergeAfterRecordMerge;
                _wordAppNo.MailMergeBeforeMergeEvent               -= OnNo_MailMergeBeforeMerge;
                _wordAppNo.MailMergeBeforeRecordMergeEvent         -= OnNo_MailMergeBeforeRecordMerge;
                _wordAppNo.MailMergeDataSourceLoadEvent            -= OnNo_MailMergeDataSourceLoad;
                _wordAppNo.MailMergeDataSourceValidateEvent        -= OnNo_MailMergeDataSourceValidate;
                _wordAppNo.MailMergeDataSourceValidate2Event       -= OnNo_MailMergeDataSourceValidate2;
                _wordAppNo.MailMergeWizardSendToCustomEvent        -= OnNo_MailMergeWizardSendToCustom;
                _wordAppNo.MailMergeWizardStateChangeEvent         -= OnNo_MailMergeWizardStateChange;
                // XML
                _wordAppNo.XMLSelectionChangeEvent                 -= OnNo_XMLSelectionChange;
                _wordAppNo.XMLValidationErrorEvent                 -= OnNo_XMLValidationError;
                // Window
                _wordAppNo.WindowActivateEvent                     -= OnNo_WindowActivate;
                _wordAppNo.WindowDeactivateEvent                   -= OnNo_WindowDeactivate;
                _wordAppNo.WindowSelectionChangeEvent              -= OnNo_WindowSelectionChange;
                _wordAppNo.WindowBeforeRightClickEvent             -= OnNo_WindowBeforeRightClick;
                _wordAppNo.WindowBeforeDoubleClickEvent            -= OnNo_WindowBeforeDoubleClick;
                _wordAppNo.WindowSizeEvent                         -= OnNo_WindowSize;
                // Protected View
                _wordAppNo.ProtectedViewWindowOpenEvent            -= OnNo_ProtectedViewWindowOpen;
                _wordAppNo.ProtectedViewWindowBeforeEditEvent      -= OnNo_ProtectedViewWindowBeforeEdit;
                _wordAppNo.ProtectedViewWindowBeforeCloseEvent     -= OnNo_ProtectedViewWindowBeforeClose;
                _wordAppNo.ProtectedViewWindowSizeEvent            -= OnNo_ProtectedViewWindowSize;
                _wordAppNo.ProtectedViewWindowActivateEvent        -= OnNo_ProtectedViewWindowActivate;
                _wordAppNo.ProtectedViewWindowDeactivateEvent      -= OnNo_ProtectedViewWindowDeactivate;
            }
            catch { }
        }
    }
}
