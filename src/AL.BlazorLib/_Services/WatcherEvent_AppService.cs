using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.AL.BlazorLib
{
    /// <summary>
    /// Shared state service that WinForms watcher forms push events into,
    /// and Blazor watcher pages subscribe to for display.
    /// </summary>
    public class WatcherEvent_AppService
    {
        /// <summary>
        /// True when the watcher pages are hosted inside an Office add-in (VSTO/COM), where the
        /// active document is expected to remain open. Add-in hosts set this to true so the
        /// "Save As JSON" UI hides the OpenXML method (which requires closing the document).
        /// </summary>
        public bool IsOfficeAddinHost { get; set; }

        // ── Word ──────────────────────────────────────────────────────────────────────

        public event Action? WordStateChanged;
        private readonly List<string> _wordLog       = new();
        private readonly List<string> _wordOpenFiles = new();
        private string? _wordActiveFile;
        private bool    _wordConnected;

        public bool    WordConnected  => _wordConnected;
        public string? WordActiveFile => _wordActiveFile;
        public IReadOnlyList<string> WordOpenFiles => _wordOpenFiles;
        public IReadOnlyList<string> WordLog       => _wordLog;

        // ── Word event toggles ────────────────────────────────────────────────────────
        private readonly List<WatcherEventToggle> _wordEventToggles = new()
        {
            new("Quit",                           "Application"),
            new("NewDocument",                    "Document"),
            new("DocumentOpen",                   "Document"),
            new("DocumentBeforeClose",            "Document"),
            new("DocumentBeforePrint",            "Document"),
            new("DocumentBeforeSave",             "Document"),
            new("DocumentChange",                 "Document"),
            new("DocumentSync",                   "Document"),
            new("EPostageInsert",                 "E-Postage"),
            new("EPostageInsertEx",               "E-Postage"),
            new("EPostagePropertyDialog",         "E-Postage"),
            new("MailMergeAfterMerge",            "Mail Merge"),
            new("MailMergeAfterRecordMerge",      "Mail Merge"),
            new("MailMergeBeforeMerge",           "Mail Merge"),
            new("MailMergeBeforeRecordMerge",     "Mail Merge"),
            new("MailMergeDataSourceLoad",        "Mail Merge"),
            new("MailMergeDataSourceValidate",    "Mail Merge"),
            new("MailMergeDataSourceValidate2",   "Mail Merge"),
            new("MailMergeWizardSendToCustom",    "Mail Merge"),
            new("MailMergeWizardStateChange",     "Mail Merge"),
            new("XMLSelectionChange",             "XML"),
            new("XMLValidationError",             "XML"),
            new("ProtectedViewWindowOpen",        "Protected View"),
            new("ProtectedViewWindowBeforeEdit",  "Protected View"),
            new("ProtectedViewWindowBeforeClose", "Protected View"),
            new("ProtectedViewWindowSize",        "Protected View"),
            new("ProtectedViewWindowActivate",    "Protected View"),
            new("ProtectedViewWindowDeactivate",  "Protected View"),
            new("WindowActivate",                 "Window"),
            new("WindowDeactivate",               "Window"),
            new("WindowSelectionChange",          "Window"),
            new("WindowBeforeRightClick",         "Window"),
            new("WindowBeforeDoubleClick",        "Window"),
            new("WindowSize",                     "Window"),
        };

        public IReadOnlyList<WatcherEventToggle> WordEventToggles => _wordEventToggles;

        public bool Word_IsEventLogged(string name)
        {
            foreach (var t in _wordEventToggles)
                if (t.Name == name) return t.Log;
            return false;
        }

        public void Word_SetEventLog(string name, bool log)
        {
            foreach (var t in _wordEventToggles)
                if (t.Name == name) { t.Log = log; WordStateChanged?.Invoke(); return; }
        }

        // ── Word delegates (set by hosting form/addin) ───────────────────────────────
        // Signature: (method, filePath, logCallback) => Task
        public Func<string, string, Action<string>, Task>? WordSaveAsJsonFunc { get; set; }

        // Called by the Blazor page to request Word activate a specific document.
        public Action<string>? WordActivateDocumentAction { get; set; }

        // Connect/disconnect — set by the hosting form; null when platform does not support it.
        public Action? WordConnectAction    { get; set; }
        public Action? WordDisconnectAction { get; set; }

        public void Word_SetConnected(bool connected, string? activeFile = null)
        {
            _wordConnected  = connected;
            _wordActiveFile = activeFile;
            WordStateChanged?.Invoke();
        }

        public void Word_UpdateOpenFiles(IEnumerable<string> files, string? activeFile)
        {
            _wordOpenFiles.Clear();
            _wordOpenFiles.AddRange(files);
            _wordActiveFile = activeFile;
            WordStateChanged?.Invoke();
        }

        public void Word_AppendLog(string message)
        {
            _wordLog.Add($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
            if (_wordLog.Count > 500) _wordLog.RemoveAt(0);
            WordStateChanged?.Invoke();
        }

        public void Word_ClearLog() { _wordLog.Clear(); WordStateChanged?.Invoke(); }

        // ── Excel ─────────────────────────────────────────────────────────────────────

        public event Action? ExcelStateChanged;
        private readonly List<string> _excelLog       = new();
        private readonly List<string> _excelOpenFiles = new();
        private string? _excelActiveFile;
        private bool _excelConnected;

        public bool   ExcelConnected  => _excelConnected;
        public string? ExcelActiveFile => _excelActiveFile;
        public IReadOnlyList<string> ExcelOpenFiles => _excelOpenFiles;
        public IReadOnlyList<string> ExcelLog       => _excelLog;

        // ── Excel event toggles ─────────────────────────────────────────────────────────
        private readonly List<WatcherEventToggle> _excelEventToggles = new()
        {
            new("AfterCalculate",                       "Calculation"),
            new("WorkbookNewChart",                     "Chart"),
            new("WorkbookRowsetComplete",               "Data"),
            new("WorkbookModelChange",                  "Data Model"),
            new("SheetPivotTableAfterValueChange",      "PivotTable"),
            new("SheetPivotTableBeforeAllocateChanges", "PivotTable"),
            new("SheetPivotTableBeforeCommitChanges",   "PivotTable"),
            new("SheetPivotTableBeforeDiscardChanges",  "PivotTable"),
            new("SheetPivotTableUpdate",                "PivotTable"),
            new("WorkbookPivotTableCloseConnection",    "PivotTable"),
            new("WorkbookPivotTableOpenConnection",     "PivotTable"),
            new("SheetTableUpdate",                     "Table"),
            new("NewWorkbook",                          "Workbook"),
            new("WorkbookActivate",                     "Workbook"),
            new("WorkbookAfterSave",                    "Workbook"),
            new("WorkbookBeforeClose",                  "Workbook"),
            new("WorkbookBeforePrint",                  "Workbook"),
            new("WorkbookBeforeSave",                   "Workbook"),
            new("WorkbookDeactivate",                   "Workbook"),
            new("WorkbookOpen",                         "Workbook"),
            new("WorkbookSync",                         "Workbook"),
            new("WorkbookAddinInstall",                 "Workbook/Add-in"),
            new("WorkbookAddinUninstall",               "Workbook/Add-in"),
            new("SheetActivate",                        "Worksheet"),
            new("SheetBeforeDelete",                    "Worksheet"),
            new("SheetBeforeDoubleClick",               "Worksheet"),
            new("SheetBeforeRightClick",                "Worksheet"),
            new("SheetCalculate",                       "Worksheet"),
            new("SheetChange",                          "Worksheet"),
            new("SheetDeactivate",                      "Worksheet"),
            new("SheetFollowHyperlink",                 "Worksheet"),
            new("SheetLensGalleryRenderComplete",       "Worksheet"),
            new("SheetSelectionChange",                 "Worksheet"),
            new("WorkbookNewSheet",                     "Worksheet"),
            new("WorkbookAfterXmlExport",               "XML"),
            new("WorkbookAfterXmlImport",               "XML"),
            new("WorkbookBeforeXmlExport",              "XML"),
            new("WorkbookBeforeXmlImport",              "XML"),
            new("ProtectedViewWindowActivate",          "Protected View"),
            new("ProtectedViewWindowBeforeClose",       "Protected View"),
            new("ProtectedViewWindowBeforeEdit",        "Protected View"),
            new("ProtectedViewWindowDeactivate",        "Protected View"),
            new("ProtectedViewWindowOpen",              "Protected View"),
            new("ProtectedViewWindowResize",            "Protected View"),
            new("WindowActivate",                       "Window"),
            new("WindowDeactivate",                     "Window"),
            new("WindowResize",                         "Window"),
        };

        public IReadOnlyList<WatcherEventToggle> ExcelEventToggles => _excelEventToggles;

        public bool Excel_IsEventLogged(string name)
        {
            foreach (var t in _excelEventToggles)
                if (t.Name == name) return t.Log;
            return false;
        }

        public void Excel_SetEventLog(string name, bool log)
        {
            foreach (var t in _excelEventToggles)
                if (t.Name == name) { t.Log = log; ExcelStateChanged?.Invoke(); return; }
        }

        // ── Excel delegates (set by hosting form/addin) ─────────────────────────────────
        // Signature: (method, filePath, logCallback) => Task
        public Func<string, string, Action<string>, Task>? ExcelSaveAsJsonFunc { get; set; }

        // Called by the Blazor page to request Excel activate a specific workbook.
        public Action<string>? ExcelActivateDocumentAction { get; set; }

        // Connect/disconnect — set by the hosting form; null when platform does not support it.
        public Action? ExcelConnectAction    { get; set; }
        public Action? ExcelDisconnectAction { get; set; }

        public void Excel_SetConnected(bool connected, string? activeFile = null)
        {
            _excelConnected  = connected;
            _excelActiveFile = activeFile;
            ExcelStateChanged?.Invoke();
        }

        public void Excel_UpdateOpenFiles(IEnumerable<string> files, string? activeFile)
        {
            _excelOpenFiles.Clear();
            _excelOpenFiles.AddRange(files);
            _excelActiveFile = activeFile;
            ExcelStateChanged?.Invoke();
        }

        public void Excel_AppendLog(string message)
        {
            _excelLog.Add($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
            if (_excelLog.Count > 500) _excelLog.RemoveAt(0);
            ExcelStateChanged?.Invoke();
        }

        public void Excel_ClearLog() { _excelLog.Clear(); ExcelStateChanged?.Invoke(); }

        // ── PowerPoint ───────────────────────────────────────────────────────────────

        public event Action? PowerPointStateChanged;
        private readonly List<string> _pptLog       = new();
        private readonly List<string> _pptOpenFiles = new();
        private string? _pptActiveFile;
        private bool _pptConnected;

        public bool   PowerPointConnected  => _pptConnected;
        public string? PowerPointActiveFile => _pptActiveFile;
        public IReadOnlyList<string> PowerPointOpenFiles => _pptOpenFiles;
        public IReadOnlyList<string> PowerPointLog       => _pptLog;

        // ── PowerPoint event toggles ────────────────────────────────────────────────────
        private readonly List<WatcherEventToggle> _pptEventToggles = new()
        {
            new("AfterNewPresentation",            "Presentation"),
            new("AfterPresentationOpen",           "Presentation"),
            new("ColorSchemeChanged",              "Presentation"),
            new("NewPresentation",                 "Presentation"),
            new("PresentationBeforeClose",         "Presentation"),
            new("PresentationBeforeSave",          "Presentation"),
            new("PresentationClose",               "Presentation"),
            new("PresentationCloseFinal",          "Presentation"),
            new("PresentationOpen",                "Presentation"),
            new("PresentationPrint",               "Presentation"),
            new("PresentationSave",                "Presentation"),
            new("PresentationSync",                "Presentation"),
            new("AfterShapeSizeChange",            "Shape"),
            new("AfterDragDropOnSlide",            "Slide"),
            new("PresentationNewSlide",            "Slide"),
            new("SlideSelectionChanged",           "Slide"),
            new("SlideShowBegin",                  "Slide Show"),
            new("SlideShowEnd",                    "Slide Show"),
            new("SlideShowNextBuild",              "Slide Show"),
            new("SlideShowNextClick",              "Slide Show"),
            new("SlideShowNextSlide",              "Slide Show"),
            new("SlideShowOnNext",                 "Slide Show"),
            new("SlideShowOnPrevious",             "Slide Show"),
            new("ProtectedViewWindowActivate",     "Protected View"),
            new("ProtectedViewWindowBeforeClose",  "Protected View"),
            new("ProtectedViewWindowBeforeEdit",   "Protected View"),
            new("ProtectedViewWindowDeactivate",   "Protected View"),
            new("ProtectedViewWindowOpen",         "Protected View"),
            new("WindowActivate",                  "Window"),
            new("WindowDeactivate",                "Window"),
            new("WindowBeforeDoubleClick",         "Window"),
            new("WindowBeforeRightClick",          "Window"),
            new("WindowSelectionChange",           "Window"),
        };

        public IReadOnlyList<WatcherEventToggle> PowerPointEventToggles => _pptEventToggles;

        public bool PowerPoint_IsEventLogged(string name)
        {
            foreach (var t in _pptEventToggles)
                if (t.Name == name) return t.Log;
            return false;
        }

        public void PowerPoint_SetEventLog(string name, bool log)
        {
            foreach (var t in _pptEventToggles)
                if (t.Name == name) { t.Log = log; PowerPointStateChanged?.Invoke(); return; }
        }

        // ── PowerPoint delegates (set by hosting form/addin) ────────────────────────────
        // Signature: (method, filePath, logCallback) => Task
        public Func<string, string, Action<string>, Task>? PowerPointSaveAsJsonFunc { get; set; }

        // Called by the Blazor page to request PowerPoint activate a specific presentation.
        public Action<string>? PowerPointActivateDocumentAction { get; set; }

        // Connect/disconnect — set by the hosting form; null when platform does not support it.
        public Action? PowerPointConnectAction    { get; set; }
        public Action? PowerPointDisconnectAction { get; set; }

        public void PowerPoint_SetConnected(bool connected, string? activeFile = null)
        {
            _pptConnected  = connected;
            _pptActiveFile = activeFile;
            PowerPointStateChanged?.Invoke();
        }

        public void PowerPoint_UpdateOpenFiles(IEnumerable<string> files, string? activeFile)
        {
            _pptOpenFiles.Clear();
            _pptOpenFiles.AddRange(files);
            _pptActiveFile = activeFile;
            PowerPointStateChanged?.Invoke();
        }

        public void PowerPoint_AppendLog(string message)
        {
            _pptLog.Add($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
            if (_pptLog.Count > 500) _pptLog.RemoveAt(0);
            PowerPointStateChanged?.Invoke();
        }

        public void PowerPoint_ClearLog() { _pptLog.Clear(); PowerPointStateChanged?.Invoke(); }
    }
}
