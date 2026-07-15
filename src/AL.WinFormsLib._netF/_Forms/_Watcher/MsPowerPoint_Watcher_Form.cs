using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using JBC.ExploreTheWorld.CL;


namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class MsPowerPoint_Watcher_Form : Form
    {
        // ── State ────────────────────────────────────────────────────────────────────
        private readonly List<MsOfficeEvent_Record> _events = MsOfficeEvents_Repo.ForPowerPoint();
        private ComboBox? _cbxDocuments;
        private bool _isConnected;
        private bool _applicationInjected;

        // Last auto-generated default output path — used so a user-entered custom path is preserved
        // when the active presentation changes.
        private string _lastDefaultOutputPath = string.Empty;

        // NetOffice
        private NetOffice.PowerPointApi.Application?  _pptAppNo;
        private NetOffice.PowerPointApi.Presentation? _pptPrsNo;

        // ── Constructor ──────────────────────────────────────────────────────────────
        public MsPowerPoint_Watcher_Form()
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
        private void MsPowerPoint_Watcher_Form_Load(object sender, EventArgs e)
        {
            _lastDefaultOutputPath = SaveAsJson_Helper.BuildDefaultPath(TrySafePrsFullName(), "ETW_MsPowerPoint", cbxJsonWriteMethod.SelectedItem?.ToString());
            txtOutputFilePath.Text = _lastDefaultOutputPath;
            cbxJsonWriteMethod.Items.AddRange(Enum.GetNames(typeof(JsonWriteMethod)));
            cbxJsonWriteMethod.SelectedItem = JsonWriteMethod.NetOffice.ToString();
            PopulateEventsGrid();
            if (_applicationInjected)
            {
                btnConnectDisconnect.Text    = "Disconnect";
                btnConnectDisconnect.Enabled = false;
                AppendLog($"Connected (via addin): {_pptPrsNo?.Name ?? "(no active presentation)"}");
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
        public void InjectApplication(NetOffice.PowerPointApi.Application app)
        {
            if (app == null) return;
            _pptAppNo             = app;
            _applicationInjected  = true;
            try { _pptPrsNo = _pptAppNo.ActivePresentation; } catch { }
            WireNetOfficeEvents();
            _isConnected = true;
            if (IsHandleCreated)
            {
                btnConnectDisconnect.Text    = "Disconnect";
                btnConnectDisconnect.Enabled = false;
                AppendLog($"Connected (via addin): {_pptPrsNo?.Name ?? "(no active presentation)"}");
            }
        }

        private void TryAutoConnect()
        {
            try
            {
                var raw   = WatcherComHelper.GetActiveCom("PowerPoint.Application");
                _pptAppNo = new NetOffice.PowerPointApi.Application(null, raw);
                _pptPrsNo = _pptAppNo.ActivePresentation;
                WireNetOfficeEvents();
                _isConnected              = true;
                btnConnectDisconnect.Text = "Disconnect";
                AppendLog($"Connected: {_pptPrsNo.Name}");
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
                FileName   = SaveAsJson_Helper.BuildDefaultPath(_pptPrsNo?.FullName, "ETW_MsPowerPoint", cbxJsonWriteMethod.SelectedItem?.ToString())
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
                if (quit) { try { _pptPrsNo?.Close(); } catch { } try { _pptAppNo?.Quit(); } catch { } }
                _pptPrsNo?.Dispose(); _pptPrsNo = null;
                if (!_applicationInjected) { _pptAppNo?.Dispose(); }
                _pptAppNo = null;
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

        // ── Default output path (follows the active presentation) ────────────────────
        private string? TrySafePrsFullName()
        {
            try { return _pptPrsNo?.FullName; } catch { return null; }
        }

        // Recompute the default path when the write method changes so the file name
        // reflects the method (e.g. filename-NetOffice.pptx.json).
        private void cbxJsonWriteMethod_SelectedIndexChanged(object sender, EventArgs e)
            => UpdateDefaultOutputPath();

        // Refreshes the default .json output path to follow the active presentation and the
        // selected write method (<active file name>-<method>.json). A user-entered custom
        // path is left untouched.
        private void UpdateDefaultOutputPath()
        {
            if (txtOutputFilePath.InvokeRequired)
            {
                try { txtOutputFilePath.BeginInvoke(new Action(UpdateDefaultOutputPath)); } catch { }
                return;
            }
            var newDefault = SaveAsJson_Helper.BuildDefaultPath(TrySafePrsFullName(), "ETW_MsPowerPoint", cbxJsonWriteMethod.SelectedItem?.ToString());
            var current    = txtOutputFilePath.Text.Trim();
            if (current.Length == 0 || current == _lastDefaultOutputPath)
                txtOutputFilePath.Text = newDefault;
            _lastDefaultOutputPath = newDefault;
        }

        // ── Save As JSON ─────────────────────────────────────────────────────────────
        private void btnSaveAsJson_Click(object sender, EventArgs e)
        {
            if (!_isConnected) { AppendLog("ERROR: No presentation connected."); return; }

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
                MsOfficeSaveAsJsonWriterProvider.Require().WritePowerPointJson(
                    method.ToString(), _pptPrsNo, _pptAppNo, path, AppendLog);
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
        private void MsPowerPoint_Watcher_Form_FormClosing(object sender, FormClosingEventArgs e)
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
                var selected = _pptPrsNo?.Name ?? _cbxDocuments.SelectedItem as string;
                _cbxDocuments.Items.Clear();
                if (_pptAppNo != null)
                {
                    foreach (NetOffice.PowerPointApi.Presentation prs in _pptAppNo.Presentations)
                    {
                        try { _cbxDocuments.Items.Add(prs.Name); } catch { }
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
            if (_cbxDocuments?.SelectedItem is not string prsName || _pptAppNo == null) return;
            try
            {
                foreach (NetOffice.PowerPointApi.Presentation prs in _pptAppNo.Presentations)
                {
                    if (prs.Name == prsName)
                    {
                        _pptPrsNo = prs;
                        try
                        {
                            if (prs.Windows.Count > 0)
                                prs.Windows[1].Activate();
                        }
                        catch (Exception ex)
                        {
                            AppendLog($"Could not activate presentation window: {ex.Message}");
                        }
                        UpdateDefaultOutputPath();
                        AppendLog($"Selected presentation: {prsName}");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Error selecting presentation: {ex.Message}");
            }
        }

        // ── NetOffice event handlers ─────────────────────────────────────────────────
        // Presentation
        private void OnNo_AfterNewPresentation(NetOffice.PowerPointApi.Presentation prs)
        {
            RefreshOpenDocumentsList();
            LogEvent("AfterNewPresentation", prs.Name);
        }
        private void OnNo_AfterPresentationOpen(NetOffice.PowerPointApi.Presentation prs)                                         => LogEvent("AfterPresentationOpen", prs.Name);
        private void OnNo_ColorSchemeChanged(NetOffice.PowerPointApi.SlideRange sldRange)                                         => LogEvent("ColorSchemeChanged", "");
        private void OnNo_NewPresentation(NetOffice.PowerPointApi.Presentation prs)                                               => LogEvent("NewPresentation", prs.Name);
        private void OnNo_PresentationBeforeClose(NetOffice.PowerPointApi.Presentation prs, ref bool cancel)                      => LogEvent("PresentationBeforeClose", prs.Name);
        private void OnNo_PresentationBeforeSave(NetOffice.PowerPointApi.Presentation prs, ref bool cancel)                       => LogEvent("PresentationBeforeSave", prs.Name);
        private void OnNo_PresentationClose(NetOffice.PowerPointApi.Presentation prs)                                             => LogEvent("PresentationClose", prs.Name);
        private void OnNo_PresentationCloseFinal(NetOffice.PowerPointApi.Presentation prs)                                        => LogEvent("PresentationCloseFinal", prs.Name);
        private void OnNo_PresentationOpen(NetOffice.PowerPointApi.Presentation prs)
        {
            RefreshOpenDocumentsList();
            LogEvent("PresentationOpen", prs.Name);
        }
        private void OnNo_PresentationPrint(NetOffice.PowerPointApi.Presentation prs)                                             => LogEvent("PresentationPrint", prs.Name);
        private void OnNo_PresentationSave(NetOffice.PowerPointApi.Presentation prs)                                              => LogEvent("PresentationSave", prs.Name);
        private void OnNo_PresentationSync(NetOffice.PowerPointApi.Presentation prs, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType) => LogEvent("PresentationSync", prs.Name);
        // Shape
        private void OnNo_AfterShapeSizeChange(NetOffice.PowerPointApi.Shape shp)                                                 => LogEvent("AfterShapeSizeChange", "");
        // Slide
        private void OnNo_AfterDragDropOnSlide(NetOffice.PowerPointApi.Slide sld, float x, float y)                               => LogEvent("AfterDragDropOnSlide", "");
        private void OnNo_PresentationNewSlide(NetOffice.PowerPointApi.Slide sld)                                                  => LogEvent("PresentationNewSlide", "");
        private void OnNo_SlideSelectionChanged(NetOffice.PowerPointApi.SlideRange sldRange)                                      => LogEvent("SlideSelectionChanged", "");
        // Slide Show
        private void OnNo_SlideShowBegin(NetOffice.PowerPointApi.SlideShowWindow wn)                                              => LogEvent("SlideShowBegin", "");
        private void OnNo_SlideShowEnd(NetOffice.PowerPointApi.Presentation prs)                                                  => LogEvent("SlideShowEnd", prs.Name);
        private void OnNo_SlideShowNextBuild(NetOffice.PowerPointApi.SlideShowWindow wn)                                          => LogEvent("SlideShowNextBuild", "");
        private void OnNo_SlideShowNextClick(NetOffice.PowerPointApi.SlideShowWindow wn, NetOffice.PowerPointApi.Effect nEffect)  => LogEvent("SlideShowNextClick", "");
        private void OnNo_SlideShowNextSlide(NetOffice.PowerPointApi.SlideShowWindow wn)                                          => LogEvent("SlideShowNextSlide", "");
        private void OnNo_SlideShowOnNext(NetOffice.PowerPointApi.SlideShowWindow wn)                                             => LogEvent("SlideShowOnNext", "");
        private void OnNo_SlideShowOnPrevious(NetOffice.PowerPointApi.SlideShowWindow wn)                                         => LogEvent("SlideShowOnPrevious", "");
        // Window
        private void OnNo_WindowActivate(NetOffice.PowerPointApi.Presentation prs, NetOffice.PowerPointApi.DocumentWindow wn)
        {
            _pptPrsNo = prs;
            RefreshOpenDocumentsList();
            UpdateDefaultOutputPath();
            LogEvent("WindowActivate", prs.Name);
        }
        private void OnNo_WindowDeactivate(NetOffice.PowerPointApi.Presentation prs, NetOffice.PowerPointApi.DocumentWindow wn)   => LogEvent("WindowDeactivate", prs.Name);
        private void OnNo_WindowBeforeRightClick(NetOffice.PowerPointApi.Selection sel, ref bool cancel)                          => LogEvent("WindowBeforeRightClick", "");
        private void OnNo_WindowBeforeDoubleClick(NetOffice.PowerPointApi.Selection sel, ref bool cancel)                         => LogEvent("WindowBeforeDoubleClick", "");
        private void OnNo_WindowSelectionChange(NetOffice.PowerPointApi.Selection sel)                                            => LogEvent("WindowSelectionChange", "");
        // Protected View
        private void OnNo_ProtectedViewWindowOpen(NetOffice.PowerPointApi.ProtectedViewWindow pvw)                                => LogEvent("ProtectedViewWindowOpen", "");
        private void OnNo_ProtectedViewWindowBeforeEdit(NetOffice.PowerPointApi.ProtectedViewWindow pvw, ref bool cancel)         => LogEvent("ProtectedViewWindowBeforeEdit", "");
        private void OnNo_ProtectedViewWindowBeforeClose(NetOffice.PowerPointApi.ProtectedViewWindow pvw, NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason closeReason, ref bool cancel) => LogEvent("ProtectedViewWindowBeforeClose", "");
        private void OnNo_ProtectedViewWindowActivate(NetOffice.PowerPointApi.ProtectedViewWindow pvw)                            => LogEvent("ProtectedViewWindowActivate", "");
        private void OnNo_ProtectedViewWindowDeactivate(NetOffice.PowerPointApi.ProtectedViewWindow pvw)                          => LogEvent("ProtectedViewWindowDeactivate", "");

        private void WireNetOfficeEvents()
        {
            if (_pptAppNo == null) return;
            // Presentation
            _pptAppNo.AfterNewPresentationEvent            += OnNo_AfterNewPresentation;
            _pptAppNo.AfterPresentationOpenEvent           += OnNo_AfterPresentationOpen;
            _pptAppNo.ColorSchemeChangedEvent              += OnNo_ColorSchemeChanged;
            _pptAppNo.NewPresentationEvent                 += OnNo_NewPresentation;
            _pptAppNo.PresentationBeforeCloseEvent         += OnNo_PresentationBeforeClose;
            _pptAppNo.PresentationBeforeSaveEvent          += OnNo_PresentationBeforeSave;
            _pptAppNo.PresentationCloseEvent               += OnNo_PresentationClose;
            _pptAppNo.PresentationCloseFinalEvent          += OnNo_PresentationCloseFinal;
            _pptAppNo.PresentationOpenEvent                += OnNo_PresentationOpen;
            _pptAppNo.PresentationPrintEvent               += OnNo_PresentationPrint;
            _pptAppNo.PresentationSaveEvent                += OnNo_PresentationSave;
            _pptAppNo.PresentationSyncEvent                += OnNo_PresentationSync;
            // Shape
            _pptAppNo.AfterShapeSizeChangeEvent            += OnNo_AfterShapeSizeChange;
            // Slide
            _pptAppNo.AfterDragDropOnSlideEvent            += OnNo_AfterDragDropOnSlide;
            _pptAppNo.PresentationNewSlideEvent            += OnNo_PresentationNewSlide;
            _pptAppNo.SlideSelectionChangedEvent           += OnNo_SlideSelectionChanged;
            // Slide Show
            _pptAppNo.SlideShowBeginEvent                  += OnNo_SlideShowBegin;
            _pptAppNo.SlideShowEndEvent                    += OnNo_SlideShowEnd;
            _pptAppNo.SlideShowNextBuildEvent              += OnNo_SlideShowNextBuild;
            _pptAppNo.SlideShowNextClickEvent              += OnNo_SlideShowNextClick;
            _pptAppNo.SlideShowNextSlideEvent              += OnNo_SlideShowNextSlide;
            _pptAppNo.SlideShowOnNextEvent                 += OnNo_SlideShowOnNext;
            _pptAppNo.SlideShowOnPreviousEvent             += OnNo_SlideShowOnPrevious;
            // Window
            _pptAppNo.WindowActivateEvent                  += OnNo_WindowActivate;
            _pptAppNo.WindowDeactivateEvent                += OnNo_WindowDeactivate;
            _pptAppNo.WindowBeforeRightClickEvent          += OnNo_WindowBeforeRightClick;
            _pptAppNo.WindowBeforeDoubleClickEvent         += OnNo_WindowBeforeDoubleClick;
            _pptAppNo.WindowSelectionChangeEvent           += OnNo_WindowSelectionChange;
            // Protected View
            _pptAppNo.ProtectedViewWindowOpenEvent         += OnNo_ProtectedViewWindowOpen;
            _pptAppNo.ProtectedViewWindowBeforeEditEvent   += OnNo_ProtectedViewWindowBeforeEdit;
            _pptAppNo.ProtectedViewWindowBeforeCloseEvent  += OnNo_ProtectedViewWindowBeforeClose;
            _pptAppNo.ProtectedViewWindowActivateEvent     += OnNo_ProtectedViewWindowActivate;
            _pptAppNo.ProtectedViewWindowDeactivateEvent   += OnNo_ProtectedViewWindowDeactivate;
        }

        private void UnwireNetOfficeEvents()
        {
            if (_pptAppNo == null) return;
            try
            {
                // Presentation
                _pptAppNo.AfterNewPresentationEvent            -= OnNo_AfterNewPresentation;
                _pptAppNo.AfterPresentationOpenEvent           -= OnNo_AfterPresentationOpen;
                _pptAppNo.ColorSchemeChangedEvent              -= OnNo_ColorSchemeChanged;
                _pptAppNo.NewPresentationEvent                 -= OnNo_NewPresentation;
                _pptAppNo.PresentationBeforeCloseEvent         -= OnNo_PresentationBeforeClose;
                _pptAppNo.PresentationBeforeSaveEvent          -= OnNo_PresentationBeforeSave;
                _pptAppNo.PresentationCloseEvent               -= OnNo_PresentationClose;
                _pptAppNo.PresentationCloseFinalEvent          -= OnNo_PresentationCloseFinal;
                _pptAppNo.PresentationOpenEvent                -= OnNo_PresentationOpen;
                _pptAppNo.PresentationPrintEvent               -= OnNo_PresentationPrint;
                _pptAppNo.PresentationSaveEvent                -= OnNo_PresentationSave;
                _pptAppNo.PresentationSyncEvent                -= OnNo_PresentationSync;
                // Shape
                _pptAppNo.AfterShapeSizeChangeEvent            -= OnNo_AfterShapeSizeChange;
                // Slide
                _pptAppNo.AfterDragDropOnSlideEvent            -= OnNo_AfterDragDropOnSlide;
                _pptAppNo.PresentationNewSlideEvent            -= OnNo_PresentationNewSlide;
                _pptAppNo.SlideSelectionChangedEvent           -= OnNo_SlideSelectionChanged;
                // Slide Show
                _pptAppNo.SlideShowBeginEvent                  -= OnNo_SlideShowBegin;
                _pptAppNo.SlideShowEndEvent                    -= OnNo_SlideShowEnd;
                _pptAppNo.SlideShowNextBuildEvent              -= OnNo_SlideShowNextBuild;
                _pptAppNo.SlideShowNextClickEvent              -= OnNo_SlideShowNextClick;
                _pptAppNo.SlideShowNextSlideEvent              -= OnNo_SlideShowNextSlide;
                _pptAppNo.SlideShowOnNextEvent                 -= OnNo_SlideShowOnNext;
                _pptAppNo.SlideShowOnPreviousEvent             -= OnNo_SlideShowOnPrevious;
                // Window
                _pptAppNo.WindowActivateEvent                  -= OnNo_WindowActivate;
                _pptAppNo.WindowDeactivateEvent                -= OnNo_WindowDeactivate;
                _pptAppNo.WindowBeforeRightClickEvent          -= OnNo_WindowBeforeRightClick;
                _pptAppNo.WindowBeforeDoubleClickEvent         -= OnNo_WindowBeforeDoubleClick;
                _pptAppNo.WindowSelectionChangeEvent           -= OnNo_WindowSelectionChange;
                // Protected View
                _pptAppNo.ProtectedViewWindowOpenEvent         -= OnNo_ProtectedViewWindowOpen;
                _pptAppNo.ProtectedViewWindowBeforeEditEvent   -= OnNo_ProtectedViewWindowBeforeEdit;
                _pptAppNo.ProtectedViewWindowBeforeCloseEvent  -= OnNo_ProtectedViewWindowBeforeClose;
                _pptAppNo.ProtectedViewWindowActivateEvent     -= OnNo_ProtectedViewWindowActivate;
                _pptAppNo.ProtectedViewWindowDeactivateEvent   -= OnNo_ProtectedViewWindowDeactivate;
            }
            catch { }
        }
    }
}
