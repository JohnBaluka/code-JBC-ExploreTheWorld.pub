using System;
using System.Collections.Generic;
using System.Windows.Forms;
using JBC.ExploreTheWorld.AL.BlazorLib;
using Microsoft.AspNetCore.Components.WebView.WindowsForms;
using Microsoft.Extensions.DependencyInjection;
using PowerPoint = NetOffice.PowerPointApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class MsPowerPoint_Watcher_WebView_Form : Form
    {
        private readonly WatcherEvent_AppService _watcherSvc;
        private PowerPoint.Application? _pptApp;

        public MsPowerPoint_Watcher_WebView_Form(ServiceProvider serviceProvider)
        {
            InitializeComponent();
            // WebView2 default UserDataFolder is derived from the host executable
            // (e.g. POWERPNT.EXE in C:\Program Files\...) which is not writable.
            // Set an explicit writable path to avoid E_ACCESSDENIED.
            blazorWebView.BlazorWebViewInitializing += (_, e) =>
                e.UserDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "JBC.ExploreTheWorld", "WebView2");
            blazorWebView.HostPage  = "wwwroot/index.html";
            blazorWebView.Services  = serviceProvider;
            blazorWebView.RootComponents.Add(
                new RootComponent("#app", typeof(Routes), null));
            blazorWebView.StartPath = "/watcher-powerpoint";

            _watcherSvc = serviceProvider.GetRequiredService<WatcherEvent_AppService>();
            Load       += OnFormLoad;
            FormClosed += OnFormClosed;
        }

        private void OnFormLoad(object? sender, EventArgs e)
        {
            try
            {
                var raw = WatcherComHelper.GetActiveCom("PowerPoint.Application");
                _pptApp = new PowerPoint.Application(null, raw);

                string? active = null;
                try { active = _pptApp.ActivePresentation?.Name; } catch { }

                WireEvents();

                _watcherSvc.PowerPoint_AppendLog("Connected to running PowerPoint instance.");
                _watcherSvc.PowerPoint_SetConnected(true, active);
                RefreshDocs();

                // Register delegates used by the Blazor page
                _watcherSvc.PowerPointSaveAsJsonFunc         = SaveAsJsonAsync;
                _watcherSvc.PowerPointActivateDocumentAction = ActivateDocument;
            }
            catch (Exception ex)
            {
                _watcherSvc.PowerPoint_AppendLog($"PowerPoint not detected: {ex.Message}");
                _watcherSvc.PowerPoint_SetConnected(false);
            }
        }

        private void OnFormClosed(object? sender, FormClosedEventArgs e)
        {
            try { UnwireEvents(); } catch { }
            try { _pptApp?.Dispose(); } catch { }
            _pptApp = null;
            _watcherSvc.PowerPointSaveAsJsonFunc         = null;
            _watcherSvc.PowerPointActivateDocumentAction = null;
            _watcherSvc.PowerPoint_SetConnected(false);
        }

        private void ActivateDocument(string prsName)
        {
            if (_pptApp == null) return;
            try
            {
                foreach (PowerPoint.Presentation prs in _pptApp.Presentations)
                {
                    try
                    {
                        if (string.Equals(prs.Name, prsName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (prs.Windows.Count > 0) prs.Windows[1].Activate();
                            return;
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex) { _watcherSvc.PowerPoint_AppendLog($"ActivateDocument: {ex.Message}"); }
        }

        // ── Log helper ───────────────────────────────────────────────────────────────

        private void LogEvent(string name, string detail = "")
        {
            if (_watcherSvc.PowerPoint_IsEventLogged(name))
                _watcherSvc.PowerPoint_AppendLog($"[Event] {name}{(detail.Length > 0 ? ": " + detail : "")}");
        }

        // Returns the current selection's text (trimmed) for event logging, e.g. "Aruba".
        private static string SelectionText(PowerPoint.Selection sel)
        {
            try
            {
                var text = sel.TextRange?.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    text = text.Replace("\r", " ").Replace("\n", " ").Trim();
                    return text.Length > 60 ? text.Substring(0, 60) + "…" : text;
                }
            }
            catch { }
            try { return sel.Type.ToString(); } catch { return string.Empty; }
        }

        // ── Save As JSON ─────────────────────────────────────────────────────────────

        private System.Threading.Tasks.Task SaveAsJsonAsync(string method, string filePath, Action<string> log)
        {
            return System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    MsOfficeSaveAsJsonWriterProvider.Require().WritePowerPointJson(
                        method, _pptApp?.ActivePresentation, _pptApp, filePath, log);
                }
                catch (Exception ex) { log($"✘ {ex.Message}"); }
            });
        }

        // ── NetOffice event wiring ────────────────────────────────────────────────────

        private void WireEvents()
        {
            if (_pptApp == null) return;
            // RefreshDocs() makes COM calls back into PowerPoint's STA; calling it
            // synchronously inside an event handler can deadlock in Office hosts.
            // Defer it with BeginInvoke so PowerPoint's STA has already unblocked.
            // Presentation
            _pptApp.AfterNewPresentationEvent           += (PowerPoint.Presentation prs) => { LogEvent("AfterNewPresentation", prs.Name); BeginInvoke(() => RefreshDocs()); };
            _pptApp.AfterPresentationOpenEvent          += (PowerPoint.Presentation prs) => { LogEvent("AfterPresentationOpen", prs.Name); BeginInvoke(() => RefreshDocs()); };
            _pptApp.ColorSchemeChangedEvent             += (PowerPoint.SlideRange sldRange) => LogEvent("ColorSchemeChanged");
            _pptApp.NewPresentationEvent                += (PowerPoint.Presentation prs) => LogEvent("NewPresentation", prs.Name);
            _pptApp.PresentationBeforeCloseEvent        += (PowerPoint.Presentation prs, ref bool c) => LogEvent("PresentationBeforeClose", prs.Name);
            _pptApp.PresentationBeforeSaveEvent         += (PowerPoint.Presentation prs, ref bool c) => LogEvent("PresentationBeforeSave", prs.Name);
            _pptApp.PresentationCloseEvent              += (PowerPoint.Presentation prs) => { LogEvent("PresentationClose", prs.Name); BeginInvoke(() => RefreshDocs()); };
            _pptApp.PresentationCloseFinalEvent         += (PowerPoint.Presentation prs) => LogEvent("PresentationCloseFinal", prs.Name);
            _pptApp.PresentationOpenEvent               += (PowerPoint.Presentation prs) => { LogEvent("PresentationOpen", prs.Name); BeginInvoke(() => RefreshDocs()); };
            _pptApp.PresentationPrintEvent              += (PowerPoint.Presentation prs) => LogEvent("PresentationPrint", prs.Name);
            _pptApp.PresentationSaveEvent               += (PowerPoint.Presentation prs) => LogEvent("PresentationSave", prs.Name);
            _pptApp.PresentationSyncEvent               += (PowerPoint.Presentation prs, NetOffice.OfficeApi.Enums.MsoSyncEventType t) => LogEvent("PresentationSync", prs.Name);
            // Shape
            _pptApp.AfterShapeSizeChangeEvent           += (PowerPoint.Shape shp) => LogEvent("AfterShapeSizeChange");
            // Slide
            _pptApp.AfterDragDropOnSlideEvent           += (PowerPoint.Slide sld, float x, float y) => LogEvent("AfterDragDropOnSlide");
            _pptApp.PresentationNewSlideEvent           += (PowerPoint.Slide sld) => LogEvent("PresentationNewSlide");
            _pptApp.SlideSelectionChangedEvent          += (PowerPoint.SlideRange sldRange) => LogEvent("SlideSelectionChanged");
            // Slide Show
            _pptApp.SlideShowBeginEvent                 += (PowerPoint.SlideShowWindow wn) => LogEvent("SlideShowBegin");
            _pptApp.SlideShowEndEvent                   += (PowerPoint.Presentation prs) => LogEvent("SlideShowEnd", prs.Name);
            _pptApp.SlideShowNextBuildEvent             += (PowerPoint.SlideShowWindow wn) => LogEvent("SlideShowNextBuild");
            _pptApp.SlideShowNextClickEvent             += (PowerPoint.SlideShowWindow wn, PowerPoint.Effect nEffect) => LogEvent("SlideShowNextClick");
            _pptApp.SlideShowNextSlideEvent             += (PowerPoint.SlideShowWindow wn) => LogEvent("SlideShowNextSlide");
            _pptApp.SlideShowOnNextEvent                += (PowerPoint.SlideShowWindow wn) => LogEvent("SlideShowOnNext");
            _pptApp.SlideShowOnPreviousEvent            += (PowerPoint.SlideShowWindow wn) => LogEvent("SlideShowOnPrevious");
            // Window
            _pptApp.WindowActivateEvent                 += (PowerPoint.Presentation prs, PowerPoint.DocumentWindow wn) => { LogEvent("WindowActivate", prs.Name); BeginInvoke(() => RefreshDocs()); };
            _pptApp.WindowDeactivateEvent               += (PowerPoint.Presentation prs, PowerPoint.DocumentWindow wn) => LogEvent("WindowDeactivate", prs.Name);
            _pptApp.WindowBeforeRightClickEvent         += (PowerPoint.Selection sel, ref bool c) => LogEvent("WindowBeforeRightClick");
            _pptApp.WindowBeforeDoubleClickEvent        += (PowerPoint.Selection sel, ref bool c) => LogEvent("WindowBeforeDoubleClick");
            _pptApp.WindowSelectionChangeEvent          += (PowerPoint.Selection sel) => LogEvent("WindowSelectionChange", SelectionText(sel));
            // Protected View
            _pptApp.ProtectedViewWindowOpenEvent        += (PowerPoint.ProtectedViewWindow pvw) => LogEvent("ProtectedViewWindowOpen");
            _pptApp.ProtectedViewWindowBeforeEditEvent  += (PowerPoint.ProtectedViewWindow pvw, ref bool c) => LogEvent("ProtectedViewWindowBeforeEdit");
            _pptApp.ProtectedViewWindowBeforeCloseEvent += (PowerPoint.ProtectedViewWindow pvw, PowerPoint.Enums.PpProtectedViewCloseReason r, ref bool c) => LogEvent("ProtectedViewWindowBeforeClose");
            _pptApp.ProtectedViewWindowActivateEvent    += (PowerPoint.ProtectedViewWindow pvw) => LogEvent("ProtectedViewWindowActivate");
            _pptApp.ProtectedViewWindowDeactivateEvent  += (PowerPoint.ProtectedViewWindow pvw) => LogEvent("ProtectedViewWindowDeactivate");
        }

        private void UnwireEvents()
        {
            // NetOffice unsubscribes all events automatically when the COM wrapper is disposed.
        }

        private void RefreshDocs()
        {
            if (_pptApp == null) return;
            var files = new List<string>();
            string? active = null;
            try
            {
                foreach (PowerPoint.Presentation prs in _pptApp.Presentations)
                    try { files.Add(prs.Name); } catch { }
                try { active = _pptApp.ActivePresentation?.Name; } catch { }
            }
            catch { }
            _watcherSvc.PowerPoint_UpdateOpenFiles(files, active);
        }
    }
}
