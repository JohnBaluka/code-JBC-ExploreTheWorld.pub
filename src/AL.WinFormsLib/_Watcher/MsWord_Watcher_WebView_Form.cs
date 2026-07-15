using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using JBC.ExploreTheWorld.AL.BlazorLib;
using Microsoft.AspNetCore.Components.WebView.WindowsForms;
using Microsoft.Extensions.DependencyInjection;
using Word = NetOffice.WordApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class MsWord_Watcher_WebView_Form : Form
    {
        private readonly WatcherEvent_AppService _watcherSvc;
        private Word.Application? _wordApp;

        public MsWord_Watcher_WebView_Form(ServiceProvider serviceProvider)
        {
            InitializeComponent();
            blazorWebView.BlazorWebViewInitializing += (_, e) =>
                e.UserDataFolder = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "JBC.ExploreTheWorld", "WebView2");
            // Absolute path so BlazorWebView resolves correctly in COM-hosted (VSTO) context.
            blazorWebView.HostPage  = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!,
                "wwwroot", "index.html");
            blazorWebView.Services  = serviceProvider;
            blazorWebView.RootComponents.Add(
                new RootComponent("#app", typeof(Routes), null));
            blazorWebView.StartPath = "/watcher-word";

            _watcherSvc = serviceProvider.GetRequiredService<WatcherEvent_AppService>();
            Load       += OnFormLoad;
            FormClosed += OnFormClosed;
        }

        private void OnFormLoad(object? sender, EventArgs e)
        {
            try
            {
                var raw    = WatcherComHelper.GetActiveCom("Word.Application");
                _wordApp   = new Word.Application(null, raw);

                var openFiles = new List<string>();
                string? activeDoc = null;
                try
                {
                    foreach (Word.Document doc in _wordApp.Documents)
                        try { openFiles.Add(doc.Name); } catch { }
                    try { activeDoc = _wordApp.ActiveDocument?.Name; } catch { }
                }
                catch { }

                WireEvents();

                _watcherSvc.Word_AppendLog("Connected to running Word instance.");
                _watcherSvc.Word_SetConnected(true, activeDoc);
                _watcherSvc.Word_UpdateOpenFiles(openFiles, activeDoc);

                // Register delegates used by the Blazor page
                _watcherSvc.WordSaveAsJsonFunc          = SaveAsJsonAsync;
                _watcherSvc.WordActivateDocumentAction  = ActivateDocument;
            }
            catch (Exception ex)
            {
                _watcherSvc.Word_AppendLog($"Word not detected: {ex.Message}");
                _watcherSvc.Word_SetConnected(false);
            }
        }

        private void OnFormClosed(object? sender, FormClosedEventArgs e)
        {
            try { UnwireEvents(); } catch { }
            try { _wordApp?.Dispose(); } catch { }
            _wordApp = null;
            _watcherSvc.WordSaveAsJsonFunc         = null;
            _watcherSvc.WordActivateDocumentAction = null;
            _watcherSvc.Word_SetConnected(false);
        }

        private void ActivateDocument(string docName)
        {
            if (_wordApp == null) return;
            try
            {
                foreach (Word.Document doc in _wordApp.Documents)
                {
                    try
                    {
                        if (string.Equals(doc.Name, docName, StringComparison.OrdinalIgnoreCase))
                        {
                            doc.Activate();
                            return;
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex) { _watcherSvc.Word_AppendLog($"ActivateDocument: {ex.Message}"); }
        }

        // ── Log helper ───────────────────────────────────────────────────────────────

        private void LogEvent(string name, string detail = "")
        {
            if (_watcherSvc.Word_IsEventLogged(name))
                _watcherSvc.Word_AppendLog($"[Event] {name}{(detail.Length > 0 ? ": " + detail : "")}");
        }

        // Returns the current selection text (trimmed) for event logging, e.g. "Aruba".
        private static string SelectionText(Word.Selection sel)
        {
            try
            {
                var text = (sel.Text ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
                return text.Length > 60 ? text.Substring(0, 60) + "…" : text;
            }
            catch { return string.Empty; }
        }

        // ── Save As JSON ─────────────────────────────────────────────────────────────

        private System.Threading.Tasks.Task SaveAsJsonAsync(string method, string filePath, Action<string> log)
        {
            return System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    MsOfficeSaveAsJsonWriterProvider.Require().WriteWordJson(
                        method, _wordApp?.ActiveDocument, _wordApp, filePath, log);
                }
                catch (Exception ex) { log($"✘ {ex.Message}"); }
            });
        }

        // ── NetOffice event wiring ────────────────────────────────────────────────────

        private void WireEvents()
        {
            if (_wordApp == null) return;
            _wordApp.QuitEvent                               += () => LogEvent("Quit");
            // RefreshDocs() makes COM calls back into Word's STA; calling it synchronously
            // inside an event handler causes COM re-entry while Word is blocking, which
            // deadlocks in Office hosts. Use BeginInvoke to defer the call to the next
            // message-pump iteration so Word's STA has already unblocked.
            _wordApp.NewDocumentEvent                        += (Word.Document doc) => { LogEvent("NewDocument", doc.Name); BeginInvoke(() => RefreshDocs()); };
            _wordApp.DocumentOpenEvent                       += (Word.Document doc) => { LogEvent("DocumentOpen", doc.Name); BeginInvoke(() => RefreshDocs()); };
            _wordApp.DocumentBeforeCloseEvent                += (Word.Document doc, ref bool c) => { LogEvent("DocumentBeforeClose", doc.Name); BeginInvoke(() => RefreshDocs()); };
            _wordApp.DocumentBeforePrintEvent                += (Word.Document doc, ref bool c) => LogEvent("DocumentBeforePrint", doc.Name);
            _wordApp.DocumentBeforeSaveEvent                 += (Word.Document doc, ref bool ui, ref bool c) => LogEvent("DocumentBeforeSave", doc.Name);
            _wordApp.DocumentChangeEvent                     += () => LogEvent("DocumentChange");
            _wordApp.DocumentSyncEvent                       += (Word.Document doc, NetOffice.OfficeApi.Enums.MsoSyncEventType t) => LogEvent("DocumentSync", doc.Name);
            _wordApp.EPostageInsertEvent                     += (Word.Document doc) => LogEvent("EPostageInsert", doc.Name);
            _wordApp.EPostageInsertExEvent                   += (Word.Document doc, int a, int b, int c2, int d, int e2, int f, string g, string h, bool i, ref bool j) => LogEvent("EPostageInsertEx", doc.Name);
            _wordApp.EPostagePropertyDialogEvent             += (Word.Document doc) => LogEvent("EPostagePropertyDialog", doc.Name);
            _wordApp.MailMergeAfterMergeEvent                += (Word.Document doc, Word.Document r) => LogEvent("MailMergeAfterMerge", doc.Name);
            _wordApp.MailMergeAfterRecordMergeEvent          += (Word.Document doc) => LogEvent("MailMergeAfterRecordMerge", doc.Name);
            _wordApp.MailMergeBeforeMergeEvent               += (Word.Document doc, int s, int e2, ref bool c) => LogEvent("MailMergeBeforeMerge", doc.Name);
            _wordApp.MailMergeBeforeRecordMergeEvent         += (Word.Document doc, ref bool c) => LogEvent("MailMergeBeforeRecordMerge", doc.Name);
            _wordApp.MailMergeDataSourceLoadEvent            += (Word.Document doc) => LogEvent("MailMergeDataSourceLoad", doc.Name);
            _wordApp.MailMergeDataSourceValidateEvent        += (Word.Document doc, ref bool h) => LogEvent("MailMergeDataSourceValidate", doc.Name);
            _wordApp.MailMergeDataSourceValidate2Event       += (Word.Document doc, ref bool h) => LogEvent("MailMergeDataSourceValidate2", doc.Name);
            _wordApp.MailMergeWizardSendToCustomEvent        += (Word.Document doc) => LogEvent("MailMergeWizardSendToCustom", doc.Name);
            _wordApp.MailMergeWizardStateChangeEvent         += (Word.Document doc, ref int f, ref int t, ref bool h) => LogEvent("MailMergeWizardStateChange", doc.Name);
            _wordApp.XMLSelectionChangeEvent                 += (Word.Selection sel, Word.XMLNode o, Word.XMLNode n, ref int r) => LogEvent("XMLSelectionChange");
            _wordApp.XMLValidationErrorEvent                 += (Word.XMLNode node) => LogEvent("XMLValidationError");
            _wordApp.WindowActivateEvent                     += (Word.Document doc, Word.Window wn) => { LogEvent("WindowActivate", doc.Name); BeginInvoke(() => RefreshDocs()); };
            _wordApp.WindowDeactivateEvent                   += (Word.Document doc, Word.Window wn) => LogEvent("WindowDeactivate", doc.Name);
            _wordApp.WindowSelectionChangeEvent              += (Word.Selection sel) => LogEvent("WindowSelectionChange", SelectionText(sel));
            _wordApp.WindowBeforeRightClickEvent             += (Word.Selection sel, ref bool c) => LogEvent("WindowBeforeRightClick");
            _wordApp.WindowBeforeDoubleClickEvent            += (Word.Selection sel, ref bool c) => LogEvent("WindowBeforeDoubleClick");
            _wordApp.WindowSizeEvent                         += (Word.Document doc, Word.Window wn) => LogEvent("WindowSize", doc.Name);
            _wordApp.ProtectedViewWindowOpenEvent            += (Word.ProtectedViewWindow pv) => LogEvent("ProtectedViewWindowOpen");
            _wordApp.ProtectedViewWindowBeforeEditEvent      += (Word.ProtectedViewWindow pv, ref bool c) => LogEvent("ProtectedViewWindowBeforeEdit");
            _wordApp.ProtectedViewWindowBeforeCloseEvent     += (Word.ProtectedViewWindow pv, int r, ref bool c) => LogEvent("ProtectedViewWindowBeforeClose");
            _wordApp.ProtectedViewWindowSizeEvent            += (Word.ProtectedViewWindow pv) => LogEvent("ProtectedViewWindowSize");
            _wordApp.ProtectedViewWindowActivateEvent        += (Word.ProtectedViewWindow pv) => LogEvent("ProtectedViewWindowActivate");
            _wordApp.ProtectedViewWindowDeactivateEvent      += (Word.ProtectedViewWindow pv) => LogEvent("ProtectedViewWindowDeactivate");
        }

        private void UnwireEvents()
        {
            // NetOffice unsubscribes all events automatically when the COM wrapper is disposed.
        }

        private void RefreshDocs()
        {
            if (_wordApp == null) return;
            var files = new List<string>();
            string? active = null;
            try
            {
                foreach (Word.Document doc in _wordApp.Documents)
                    try { files.Add(doc.Name); } catch { }
                try { active = _wordApp.ActiveDocument?.Name; } catch { }
            }
            catch { }
            _watcherSvc.Word_UpdateOpenFiles(files, active);
        }
    }
}
