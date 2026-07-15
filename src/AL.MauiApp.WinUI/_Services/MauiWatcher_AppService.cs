using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.AL.WinFormsLib;
using Word = NetOffice.WordApi;
using Excel = NetOffice.ExcelApi;
using PowerPoint = NetOffice.PowerPointApi;

namespace JBC.ExploreTheWorld.AL.MauiApp.WinUI;

/// <summary>
/// Windows-head watcher bridge: connects the Blazor watcher pages (via the shared
/// <see cref="WatcherEvent_AppService"/>) to running Word/Excel/PowerPoint instances through
/// NetOffice COM, mirroring <c>AL.WinFormsLib.ExploreTheWorld_Form</c>. COM connect/disconnect
/// is marshalled to the MAUI main (STA) thread via <see cref="MainThread"/>.
/// </summary>
internal sealed class MauiWatcher_AppService
{
    private readonly WatcherEvent_AppService _watcherSvc;
    private Word.Application? _wordApp;
    private Excel.Application? _excelApp;
    private PowerPoint.Application? _pptApp;

    public MauiWatcher_AppService(WatcherEvent_AppService watcherSvc)
    {
        _watcherSvc = watcherSvc;
        _watcherSvc.WordConnectAction          = ConnectToWord;
        _watcherSvc.WordDisconnectAction       = DisconnectFromWord;
        _watcherSvc.ExcelConnectAction         = ConnectToExcel;
        _watcherSvc.ExcelDisconnectAction      = DisconnectFromExcel;
        _watcherSvc.PowerPointConnectAction    = ConnectToPowerPoint;
        _watcherSvc.PowerPointDisconnectAction = DisconnectFromPowerPoint;
    }

    // Safely evaluates a COM property for event-log detail (e.g. the selected range address).
    private static string TrySafeText(Func<string> get)
    {
        try { return get() ?? string.Empty; }
        catch { return string.Empty; }
    }

    // ── Word COM connection (STA-marshalled via MainThread) ───────────────────

    private void ConnectToWord()
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            try
            {
                var raw = WatcherComHelper.GetOrCreateCom("Word.Application", out bool created);
                _wordApp = new Word.Application(null, raw);
                if (created) _wordApp.Visible = true;

                var openFiles = new List<string>();
                string? activeDoc = null;
                try
                {
                    foreach (Word.Document doc in _wordApp.Documents)
                        try { openFiles.Add(doc.Name); } catch { }
                    try { activeDoc = _wordApp.ActiveDocument?.Name; } catch { }
                }
                catch { }

                WireWordEvents();

                _watcherSvc.Word_AppendLog(created
                    ? "Started a new Word instance."
                    : "Connected to running Word instance.");
                _watcherSvc.Word_SetConnected(true, activeDoc);
                _watcherSvc.Word_UpdateOpenFiles(openFiles, activeDoc);
                _watcherSvc.WordSaveAsJsonFunc         = SaveWordAsJsonAsync;
                _watcherSvc.WordActivateDocumentAction = ActivateWordDocument;
            }
            catch (Exception ex)
            {
                _watcherSvc.Word_AppendLog($"Word not detected: {ex.Message}");
                _watcherSvc.Word_SetConnected(false);
            }
        });
    }

    private void DisconnectFromWord()
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            try { _wordApp?.Dispose(); } catch { }
            _wordApp = null;
            _watcherSvc.WordSaveAsJsonFunc         = null;
            _watcherSvc.WordActivateDocumentAction = null;
            _watcherSvc.Word_SetConnected(false);
        });
    }

    private void ActivateWordDocument(string docName)
    {
        if (_wordApp == null) return;
        try
        {
            foreach (Word.Document doc in _wordApp.Documents)
                try
                {
                    if (string.Equals(doc.Name, docName, StringComparison.OrdinalIgnoreCase))
                    { doc.Activate(); return; }
                }
                catch { }
        }
        catch (Exception ex) { _watcherSvc.Word_AppendLog($"ActivateDocument: {ex.Message}"); }
    }

    // ── Word Save As Json ─────────────────────────────────────────────────────

    private Task SaveWordAsJsonAsync(string method, string filePath, Action<string> log)
    {
        return Task.Run(() =>
        {
            try
            {
                MsOfficeSaveAsJsonWriterProvider.Require().WriteWordJson(
                    method, _wordApp?.ActiveDocument, _wordApp, filePath, log);
            }
            catch (Exception ex) { log($"✘ {ex.Message}"); }
        });
    }

    // ── Word event log helper ─────────────────────────────────────────────────

    private void LogWordEvent(string name, string detail = "")
    {
        if (_watcherSvc.Word_IsEventLogged(name))
            _watcherSvc.Word_AppendLog($"[Event] {name}{(detail.Length > 0 ? ": " + detail : "")}");
    }

    // ── Word COM event wiring ─────────────────────────────────────────────────

    private void WireWordEvents()
    {
        if (_wordApp == null) return;
        _wordApp.QuitEvent                          += () => LogWordEvent("Quit");
        _wordApp.NewDocumentEvent                   += (Word.Document doc) => { RefreshWordDocs(); LogWordEvent("NewDocument", doc.Name); };
        _wordApp.DocumentOpenEvent                  += (Word.Document doc) => { RefreshWordDocs(); LogWordEvent("DocumentOpen", doc.Name); };
        _wordApp.DocumentBeforeCloseEvent           += (Word.Document doc, ref bool c) => { RefreshWordDocs(); LogWordEvent("DocumentBeforeClose", doc.Name); };
        _wordApp.DocumentBeforePrintEvent           += (Word.Document doc, ref bool c) => LogWordEvent("DocumentBeforePrint", doc.Name);
        _wordApp.DocumentBeforeSaveEvent            += (Word.Document doc, ref bool ui, ref bool c) => LogWordEvent("DocumentBeforeSave", doc.Name);
        _wordApp.DocumentChangeEvent                += () => LogWordEvent("DocumentChange");
        _wordApp.DocumentSyncEvent                  += (Word.Document doc, NetOffice.OfficeApi.Enums.MsoSyncEventType t) => LogWordEvent("DocumentSync", doc.Name);
        _wordApp.EPostageInsertEvent                += (Word.Document doc) => LogWordEvent("EPostageInsert", doc.Name);
        _wordApp.EPostageInsertExEvent              += (Word.Document doc, int a, int b, int c2, int d, int e2, int f, string g, string h, bool i, ref bool j) => LogWordEvent("EPostageInsertEx", doc.Name);
        _wordApp.EPostagePropertyDialogEvent        += (Word.Document doc) => LogWordEvent("EPostagePropertyDialog", doc.Name);
        _wordApp.MailMergeAfterMergeEvent           += (Word.Document doc, Word.Document r) => LogWordEvent("MailMergeAfterMerge", doc.Name);
        _wordApp.MailMergeAfterRecordMergeEvent     += (Word.Document doc) => LogWordEvent("MailMergeAfterRecordMerge", doc.Name);
        _wordApp.MailMergeBeforeMergeEvent          += (Word.Document doc, int s, int e2, ref bool c) => LogWordEvent("MailMergeBeforeMerge", doc.Name);
        _wordApp.MailMergeBeforeRecordMergeEvent    += (Word.Document doc, ref bool c) => LogWordEvent("MailMergeBeforeRecordMerge", doc.Name);
        _wordApp.MailMergeDataSourceLoadEvent       += (Word.Document doc) => LogWordEvent("MailMergeDataSourceLoad", doc.Name);
        _wordApp.MailMergeDataSourceValidateEvent   += (Word.Document doc, ref bool h) => LogWordEvent("MailMergeDataSourceValidate", doc.Name);
        _wordApp.MailMergeDataSourceValidate2Event  += (Word.Document doc, ref bool h) => LogWordEvent("MailMergeDataSourceValidate2", doc.Name);
        _wordApp.MailMergeWizardSendToCustomEvent   += (Word.Document doc) => LogWordEvent("MailMergeWizardSendToCustom", doc.Name);
        _wordApp.MailMergeWizardStateChangeEvent    += (Word.Document doc, ref int f, ref int t, ref bool h) => LogWordEvent("MailMergeWizardStateChange", doc.Name);
        _wordApp.XMLSelectionChangeEvent            += (Word.Selection sel, Word.XMLNode o, Word.XMLNode n, ref int r) => LogWordEvent("XMLSelectionChange");
        _wordApp.XMLValidationErrorEvent            += (Word.XMLNode node) => LogWordEvent("XMLValidationError");
        _wordApp.WindowActivateEvent                += (Word.Document doc, Word.Window wn) => { RefreshWordDocs(); LogWordEvent("WindowActivate", doc.Name); };
        _wordApp.WindowDeactivateEvent              += (Word.Document doc, Word.Window wn) => LogWordEvent("WindowDeactivate", doc.Name);
        _wordApp.WindowSelectionChangeEvent         += (Word.Selection sel) => LogWordEvent("WindowSelectionChange");
        _wordApp.WindowBeforeRightClickEvent        += (Word.Selection sel, ref bool c) => LogWordEvent("WindowBeforeRightClick");
        _wordApp.WindowBeforeDoubleClickEvent       += (Word.Selection sel, ref bool c) => LogWordEvent("WindowBeforeDoubleClick");
        _wordApp.WindowSizeEvent                    += (Word.Document doc, Word.Window wn) => LogWordEvent("WindowSize", doc.Name);
        _wordApp.ProtectedViewWindowOpenEvent           += (Word.ProtectedViewWindow pv) => LogWordEvent("ProtectedViewWindowOpen");
        _wordApp.ProtectedViewWindowBeforeEditEvent     += (Word.ProtectedViewWindow pv, ref bool c) => LogWordEvent("ProtectedViewWindowBeforeEdit");
        _wordApp.ProtectedViewWindowBeforeCloseEvent    += (Word.ProtectedViewWindow pv, int r, ref bool c) => LogWordEvent("ProtectedViewWindowBeforeClose");
        _wordApp.ProtectedViewWindowSizeEvent           += (Word.ProtectedViewWindow pv) => LogWordEvent("ProtectedViewWindowSize");
        _wordApp.ProtectedViewWindowActivateEvent       += (Word.ProtectedViewWindow pv) => LogWordEvent("ProtectedViewWindowActivate");
        _wordApp.ProtectedViewWindowDeactivateEvent     += (Word.ProtectedViewWindow pv) => LogWordEvent("ProtectedViewWindowDeactivate");
    }

    private void RefreshWordDocs()
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

    // ── Excel COM connection (STA-marshalled via MainThread) ──────────────────

    private void ConnectToExcel()
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            try
            {
                var raw   = WatcherComHelper.GetOrCreateCom("Excel.Application", out bool created);
                _excelApp = new Excel.Application(null, raw);
                if (created) _excelApp.Visible = true;

                string? active = null;
                try { active = _excelApp.ActiveWorkbook?.Name; } catch { }

                WireExcelEvents();

                _watcherSvc.Excel_AppendLog(created
                    ? "Started a new Excel instance."
                    : "Connected to running Excel instance.");
                _watcherSvc.Excel_SetConnected(true, active);
                RefreshExcelDocs();
                _watcherSvc.ExcelSaveAsJsonFunc         = SaveExcelAsJsonAsync;
                _watcherSvc.ExcelActivateDocumentAction = ActivateExcelWorkbook;
            }
            catch (Exception ex)
            {
                _watcherSvc.Excel_AppendLog($"Excel not detected: {ex.Message}");
                _watcherSvc.Excel_SetConnected(false);
            }
        });
    }

    private void DisconnectFromExcel()
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            try { _excelApp?.Dispose(); } catch { }
            _excelApp = null;
            _watcherSvc.ExcelSaveAsJsonFunc         = null;
            _watcherSvc.ExcelActivateDocumentAction = null;
            _watcherSvc.Excel_SetConnected(false);
        });
    }

    private void ActivateExcelWorkbook(string wbName)
    {
        if (_excelApp == null) return;
        try
        {
            foreach (Excel.Workbook wb in _excelApp.Workbooks)
                try
                {
                    if (string.Equals(wb.Name, wbName, StringComparison.OrdinalIgnoreCase))
                    { wb.Activate(); return; }
                }
                catch { }
        }
        catch (Exception ex) { _watcherSvc.Excel_AppendLog($"ActivateDocument: {ex.Message}"); }
    }

    // ── Excel Save As Json ────────────────────────────────────────────────────

    private Task SaveExcelAsJsonAsync(string method, string filePath, Action<string> log)
    {
        return Task.Run(() =>
        {
            try
            {
                MsOfficeSaveAsJsonWriterProvider.Require().WriteExcelJson(
                    method, _excelApp?.ActiveWorkbook, _excelApp, filePath, log);
            }
            catch (Exception ex) { log($"✘ {ex.Message}"); }
        });
    }

    // ── Excel event log helper ────────────────────────────────────────────────

    private void LogExcelEvent(string name, string detail = "")
    {
        if (_watcherSvc.Excel_IsEventLogged(name))
            _watcherSvc.Excel_AppendLog($"[Event] {name}{(detail.Length > 0 ? ": " + detail : "")}");
    }

    // ── Excel COM event wiring ────────────────────────────────────────────────

    private void WireExcelEvents()
    {
        if (_excelApp == null) return;
        _excelApp.NewWorkbookEvent          += (Excel.Workbook wb) => { LogExcelEvent("NewWorkbook", wb.Name); MainThread.BeginInvokeOnMainThread(RefreshExcelDocs); };
        _excelApp.WorkbookActivateEvent     += (Excel.Workbook wb) => { LogExcelEvent("WorkbookActivate", wb.Name); MainThread.BeginInvokeOnMainThread(RefreshExcelDocs); };
        _excelApp.WorkbookAfterSaveEvent    += (Excel.Workbook wb, bool success) => LogExcelEvent("WorkbookAfterSave", wb.Name);
        _excelApp.WorkbookBeforeCloseEvent  += (Excel.Workbook wb, ref bool c) => { LogExcelEvent("WorkbookBeforeClose", wb.Name); MainThread.BeginInvokeOnMainThread(RefreshExcelDocs); };
        _excelApp.WorkbookBeforePrintEvent  += (Excel.Workbook wb, ref bool c) => LogExcelEvent("WorkbookBeforePrint", wb.Name);
        _excelApp.WorkbookBeforeSaveEvent   += (Excel.Workbook wb, bool saveAsUi, ref bool c) => LogExcelEvent("WorkbookBeforeSave", wb.Name);
        _excelApp.WorkbookDeactivateEvent   += (Excel.Workbook wb) => LogExcelEvent("WorkbookDeactivate", wb.Name);
        _excelApp.WorkbookOpenEvent         += (Excel.Workbook wb) => { LogExcelEvent("WorkbookOpen", wb.Name); MainThread.BeginInvokeOnMainThread(RefreshExcelDocs); };
        _excelApp.WorkbookNewSheetEvent     += (Excel.Workbook wb, NetOffice.ICOMObject sh) => LogExcelEvent("WorkbookNewSheet", wb.Name);
        _excelApp.SheetActivateEvent        += (NetOffice.ICOMObject sh) => LogExcelEvent("SheetActivate");
        _excelApp.SheetDeactivateEvent      += (NetOffice.ICOMObject sh) => LogExcelEvent("SheetDeactivate");
        _excelApp.SheetChangeEvent          += (NetOffice.ICOMObject sh, Excel.Range target) => LogExcelEvent("SheetChange", TrySafeText(() => target.Address));
        _excelApp.SheetSelectionChangeEvent += (NetOffice.ICOMObject sh, Excel.Range target) => LogExcelEvent("SheetSelectionChange", TrySafeText(() => target.Address));
        _excelApp.SheetCalculateEvent       += (NetOffice.ICOMObject sh) => LogExcelEvent("SheetCalculate");
        _excelApp.WindowActivateEvent       += (Excel.Workbook wb, Excel.Window wn) => { LogExcelEvent("WindowActivate", wb.Name); MainThread.BeginInvokeOnMainThread(RefreshExcelDocs); };
        _excelApp.WindowDeactivateEvent     += (Excel.Workbook wb, Excel.Window wn) => LogExcelEvent("WindowDeactivate", wb.Name);
        _excelApp.WindowResizeEvent         += (Excel.Workbook wb, Excel.Window wn) => LogExcelEvent("WindowResize", wb.Name);
    }

    private void RefreshExcelDocs()
    {
        if (_excelApp == null) return;
        var files = new List<string>();
        string? active = null;
        try
        {
            foreach (Excel.Workbook wb in _excelApp.Workbooks)
                try { files.Add(wb.Name); } catch { }
            try { active = _excelApp.ActiveWorkbook?.Name; } catch { }
        }
        catch { }
        _watcherSvc.Excel_UpdateOpenFiles(files, active);
    }

    // ── PowerPoint COM connection (STA-marshalled via MainThread) ─────────────

    private void ConnectToPowerPoint()
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            try
            {
                var raw = WatcherComHelper.GetOrCreateCom("PowerPoint.Application", out bool created);
                _pptApp = new PowerPoint.Application(null, raw);
                if (created) _pptApp.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;

                string? active = null;
                try { active = _pptApp.ActivePresentation?.Name; } catch { }

                WirePowerPointEvents();

                _watcherSvc.PowerPoint_AppendLog(created
                    ? "Started a new PowerPoint instance."
                    : "Connected to running PowerPoint instance.");
                _watcherSvc.PowerPoint_SetConnected(true, active);
                RefreshPowerPointDocs();
                _watcherSvc.PowerPointSaveAsJsonFunc         = SavePowerPointAsJsonAsync;
                _watcherSvc.PowerPointActivateDocumentAction = ActivatePowerPointPresentation;
            }
            catch (Exception ex)
            {
                _watcherSvc.PowerPoint_AppendLog($"PowerPoint not detected: {ex.Message}");
                _watcherSvc.PowerPoint_SetConnected(false);
            }
        });
    }

    private void DisconnectFromPowerPoint()
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            try { _pptApp?.Dispose(); } catch { }
            _pptApp = null;
            _watcherSvc.PowerPointSaveAsJsonFunc         = null;
            _watcherSvc.PowerPointActivateDocumentAction = null;
            _watcherSvc.PowerPoint_SetConnected(false);
        });
    }

    private void ActivatePowerPointPresentation(string prsName)
    {
        if (_pptApp == null) return;
        try
        {
            foreach (PowerPoint.Presentation prs in _pptApp.Presentations)
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
        catch (Exception ex) { _watcherSvc.PowerPoint_AppendLog($"ActivateDocument: {ex.Message}"); }
    }

    // ── PowerPoint Save As Json ───────────────────────────────────────────────

    private Task SavePowerPointAsJsonAsync(string method, string filePath, Action<string> log)
    {
        return Task.Run(() =>
        {
            try
            {
                MsOfficeSaveAsJsonWriterProvider.Require().WritePowerPointJson(
                    method, _pptApp?.ActivePresentation, _pptApp, filePath, log);
            }
            catch (Exception ex) { log($"✘ {ex.Message}"); }
        });
    }

    // ── PowerPoint event log helper ───────────────────────────────────────────

    private void LogPowerPointEvent(string name, string detail = "")
    {
        if (_watcherSvc.PowerPoint_IsEventLogged(name))
            _watcherSvc.PowerPoint_AppendLog($"[Event] {name}{(detail.Length > 0 ? ": " + detail : "")}");
    }

    // Returns the current selection's text (trimmed) for event logging, e.g. "Aruba".
    private static string PowerPointSelectionText(PowerPoint.Selection sel)
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

    // ── PowerPoint COM event wiring ───────────────────────────────────────────

    private void WirePowerPointEvents()
    {
        if (_pptApp == null) return;
        _pptApp.AfterNewPresentationEvent    += (PowerPoint.Presentation prs) => { LogPowerPointEvent("AfterNewPresentation", prs.Name); MainThread.BeginInvokeOnMainThread(RefreshPowerPointDocs); };
        _pptApp.AfterPresentationOpenEvent   += (PowerPoint.Presentation prs) => { LogPowerPointEvent("AfterPresentationOpen", prs.Name); MainThread.BeginInvokeOnMainThread(RefreshPowerPointDocs); };
        _pptApp.NewPresentationEvent         += (PowerPoint.Presentation prs) => LogPowerPointEvent("NewPresentation", prs.Name);
        _pptApp.PresentationBeforeCloseEvent += (PowerPoint.Presentation prs, ref bool c) => LogPowerPointEvent("PresentationBeforeClose", prs.Name);
        _pptApp.PresentationBeforeSaveEvent  += (PowerPoint.Presentation prs, ref bool c) => LogPowerPointEvent("PresentationBeforeSave", prs.Name);
        _pptApp.PresentationCloseEvent       += (PowerPoint.Presentation prs) => { LogPowerPointEvent("PresentationClose", prs.Name); MainThread.BeginInvokeOnMainThread(RefreshPowerPointDocs); };
        _pptApp.PresentationOpenEvent        += (PowerPoint.Presentation prs) => { LogPowerPointEvent("PresentationOpen", prs.Name); MainThread.BeginInvokeOnMainThread(RefreshPowerPointDocs); };
        _pptApp.PresentationSaveEvent        += (PowerPoint.Presentation prs) => LogPowerPointEvent("PresentationSave", prs.Name);
        _pptApp.SlideSelectionChangedEvent   += (PowerPoint.SlideRange sldRange) => LogPowerPointEvent("SlideSelectionChanged");
        _pptApp.SlideShowBeginEvent          += (PowerPoint.SlideShowWindow wn) => LogPowerPointEvent("SlideShowBegin");
        _pptApp.SlideShowEndEvent            += (PowerPoint.Presentation prs) => LogPowerPointEvent("SlideShowEnd", prs.Name);
        _pptApp.SlideShowNextSlideEvent      += (PowerPoint.SlideShowWindow wn) => LogPowerPointEvent("SlideShowNextSlide");
        _pptApp.WindowActivateEvent          += (PowerPoint.Presentation prs, PowerPoint.DocumentWindow wn) => { LogPowerPointEvent("WindowActivate", prs.Name); MainThread.BeginInvokeOnMainThread(RefreshPowerPointDocs); };
        _pptApp.WindowDeactivateEvent        += (PowerPoint.Presentation prs, PowerPoint.DocumentWindow wn) => LogPowerPointEvent("WindowDeactivate", prs.Name);
        _pptApp.WindowSelectionChangeEvent   += (PowerPoint.Selection sel) => LogPowerPointEvent("WindowSelectionChange", PowerPointSelectionText(sel));
    }

    private void RefreshPowerPointDocs()
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
