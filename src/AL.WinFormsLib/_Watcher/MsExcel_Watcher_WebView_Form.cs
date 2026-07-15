using System;
using System.Collections.Generic;
using System.Windows.Forms;
using JBC.ExploreTheWorld.AL.BlazorLib;
using Microsoft.AspNetCore.Components.WebView.WindowsForms;
using Microsoft.Extensions.DependencyInjection;
using Excel = NetOffice.ExcelApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class MsExcel_Watcher_WebView_Form : Form
    {
        private readonly WatcherEvent_AppService _watcherSvc;
        private Excel.Application? _excelApp;

        public MsExcel_Watcher_WebView_Form(ServiceProvider serviceProvider)
        {
            InitializeComponent();
            // WebView2 default UserDataFolder is derived from the host executable
            // (e.g. EXCEL.EXE in C:\Program Files\...) which is not writable.
            // Set an explicit writable path to avoid E_ACCESSDENIED.
            blazorWebView.BlazorWebViewInitializing += (_, e) =>
                e.UserDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "JBC.ExploreTheWorld", "WebView2");
            blazorWebView.HostPage  = "wwwroot/index.html";
            blazorWebView.Services  = serviceProvider;
            blazorWebView.RootComponents.Add(
                new RootComponent("#app", typeof(Routes), null));
            blazorWebView.StartPath = "/watcher-excel";

            _watcherSvc = serviceProvider.GetRequiredService<WatcherEvent_AppService>();
            Load       += OnFormLoad;
            FormClosed += OnFormClosed;
        }

        private void OnFormLoad(object? sender, EventArgs e)
        {
            try
            {
                var raw   = WatcherComHelper.GetActiveCom("Excel.Application");
                _excelApp = new Excel.Application(null, raw);

                string? active = null;
                try { active = _excelApp.ActiveWorkbook?.Name; } catch { }

                WireEvents();

                _watcherSvc.Excel_AppendLog("Connected to running Excel instance.");
                _watcherSvc.Excel_SetConnected(true, active);
                RefreshDocs();

                // Register delegates used by the Blazor page
                _watcherSvc.ExcelSaveAsJsonFunc         = SaveAsJsonAsync;
                _watcherSvc.ExcelActivateDocumentAction = ActivateDocument;
            }
            catch (Exception ex)
            {
                _watcherSvc.Excel_AppendLog($"Excel not detected: {ex.Message}");
                _watcherSvc.Excel_SetConnected(false);
            }
        }

        private void OnFormClosed(object? sender, FormClosedEventArgs e)
        {
            try { UnwireEvents(); } catch { }
            try { _excelApp?.Dispose(); } catch { }
            _excelApp = null;
            _watcherSvc.ExcelSaveAsJsonFunc         = null;
            _watcherSvc.ExcelActivateDocumentAction = null;
            _watcherSvc.Excel_SetConnected(false);
        }

        private void ActivateDocument(string wbName)
        {
            if (_excelApp == null) return;
            try
            {
                foreach (Excel.Workbook wb in _excelApp.Workbooks)
                {
                    try
                    {
                        if (string.Equals(wb.Name, wbName, StringComparison.OrdinalIgnoreCase))
                        {
                            wb.Activate();
                            return;
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex) { _watcherSvc.Excel_AppendLog($"ActivateDocument: {ex.Message}"); }
        }

        // ── Log helper ───────────────────────────────────────────────────────────────

        private void LogEvent(string name, string detail = "")
        {
            if (_watcherSvc.Excel_IsEventLogged(name))
                _watcherSvc.Excel_AppendLog($"[Event] {name}{(detail.Length > 0 ? ": " + detail : "")}");
        }

        // Safely evaluates a COM property for event-log detail (e.g. the selected range address).
        private static string TrySafeText(Func<string> get)
        {
            try { return get() ?? string.Empty; }
            catch { return string.Empty; }
        }

        // ── Save As JSON ─────────────────────────────────────────────────────────────

        private System.Threading.Tasks.Task SaveAsJsonAsync(string method, string filePath, Action<string> log)
        {
            return System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    MsOfficeSaveAsJsonWriterProvider.Require().WriteExcelJson(
                        method, _excelApp?.ActiveWorkbook, _excelApp, filePath, log);
                }
                catch (Exception ex) { log($"✘ {ex.Message}"); }
            });
        }

        // ── NetOffice event wiring ────────────────────────────────────────────────────

        private void WireEvents()
        {
            if (_excelApp == null) return;
            // RefreshDocs() makes COM calls back into Excel's STA; calling it
            // synchronously inside an event handler can deadlock in Office hosts.
            // Defer it with BeginInvoke so Excel's STA has already unblocked.
            // Calculation
            _excelApp.AfterCalculateEvent                       += () => LogEvent("AfterCalculate");
            // Chart
            _excelApp.WorkbookNewChartEvent                     += (Excel.Workbook wb, Excel.Chart ch) => LogEvent("WorkbookNewChart", wb.Name);
            // Data
            _excelApp.WorkbookRowsetCompleteEvent               += (Excel.Workbook wb, string desc, string sheet, bool success) => LogEvent("WorkbookRowsetComplete", wb.Name);
            // Data Model
            _excelApp.WorkbookModelChangeEvent                  += (Excel.Workbook wb, Excel.ModelChanges changes) => LogEvent("WorkbookModelChange", wb.Name);
            // PivotTable
            _excelApp.SheetPivotTableAfterValueChangeEvent      += (NetOffice.ICOMObject sh, Excel.PivotTable pt, Excel.Range tr) => LogEvent("SheetPivotTableAfterValueChange");
            _excelApp.SheetPivotTableBeforeAllocateChangesEvent += (NetOffice.ICOMObject sh, Excel.PivotTable pt, int s, int e2, ref bool c) => LogEvent("SheetPivotTableBeforeAllocateChanges");
            _excelApp.SheetPivotTableBeforeCommitChangesEvent   += (NetOffice.ICOMObject sh, Excel.PivotTable pt, int s, int e2, ref bool c) => LogEvent("SheetPivotTableBeforeCommitChanges");
            _excelApp.SheetPivotTableBeforeDiscardChangesEvent  += (NetOffice.ICOMObject sh, Excel.PivotTable pt, int s, int e2) => LogEvent("SheetPivotTableBeforeDiscardChanges");
            _excelApp.SheetPivotTableUpdateEvent                += (NetOffice.ICOMObject sh, Excel.PivotTable target) => LogEvent("SheetPivotTableUpdate");
            _excelApp.WorkbookPivotTableCloseConnectionEvent    += (Excel.Workbook wb, Excel.PivotTable target) => LogEvent("WorkbookPivotTableCloseConnection", wb.Name);
            _excelApp.WorkbookPivotTableOpenConnectionEvent     += (Excel.Workbook wb, Excel.PivotTable target) => LogEvent("WorkbookPivotTableOpenConnection", wb.Name);
            // Table
            _excelApp.SheetTableUpdateEvent                     += (NetOffice.ICOMObject sh, Excel.TableObject target) => LogEvent("SheetTableUpdate");
            // Workbook
            _excelApp.NewWorkbookEvent                          += (Excel.Workbook wb) => { LogEvent("NewWorkbook", wb.Name); BeginInvoke(() => RefreshDocs()); };
            _excelApp.WorkbookActivateEvent                     += (Excel.Workbook wb) => { LogEvent("WorkbookActivate", wb.Name); BeginInvoke(() => RefreshDocs()); };
            _excelApp.WorkbookAfterSaveEvent                    += (Excel.Workbook wb, bool success) => LogEvent("WorkbookAfterSave", wb.Name);
            _excelApp.WorkbookBeforeCloseEvent                  += (Excel.Workbook wb, ref bool c) => { LogEvent("WorkbookBeforeClose", wb.Name); BeginInvoke(() => RefreshDocs()); };
            _excelApp.WorkbookBeforePrintEvent                  += (Excel.Workbook wb, ref bool c) => LogEvent("WorkbookBeforePrint", wb.Name);
            _excelApp.WorkbookBeforeSaveEvent                   += (Excel.Workbook wb, bool saveAsUi, ref bool c) => LogEvent("WorkbookBeforeSave", wb.Name);
            _excelApp.WorkbookDeactivateEvent                   += (Excel.Workbook wb) => LogEvent("WorkbookDeactivate", wb.Name);
            _excelApp.WorkbookOpenEvent                         += (Excel.Workbook wb) => { LogEvent("WorkbookOpen", wb.Name); BeginInvoke(() => RefreshDocs()); };
            _excelApp.WorkbookSyncEvent                         += (Excel.Workbook wb, NetOffice.OfficeApi.Enums.MsoSyncEventType t) => LogEvent("WorkbookSync", wb.Name);
            // Workbook/Add-in
            _excelApp.WorkbookAddinInstallEvent                 += (Excel.Workbook wb) => LogEvent("WorkbookAddinInstall", wb.Name);
            _excelApp.WorkbookAddinUninstallEvent               += (Excel.Workbook wb) => LogEvent("WorkbookAddinUninstall", wb.Name);
            // Worksheet
            _excelApp.SheetActivateEvent                        += (NetOffice.ICOMObject sh) => LogEvent("SheetActivate");
            _excelApp.SheetBeforeDeleteEvent                    += (NetOffice.ICOMObject sh) => LogEvent("SheetBeforeDelete");
            _excelApp.SheetBeforeDoubleClickEvent               += (NetOffice.ICOMObject sh, Excel.Range target, ref bool c) => LogEvent("SheetBeforeDoubleClick");
            _excelApp.SheetBeforeRightClickEvent                += (NetOffice.ICOMObject sh, Excel.Range target, ref bool c) => LogEvent("SheetBeforeRightClick");
            _excelApp.SheetCalculateEvent                       += (NetOffice.ICOMObject sh) => LogEvent("SheetCalculate");
            _excelApp.SheetChangeEvent                          += (NetOffice.ICOMObject sh, Excel.Range target) => LogEvent("SheetChange", target.Address);
            _excelApp.SheetDeactivateEvent                      += (NetOffice.ICOMObject sh) => LogEvent("SheetDeactivate");
            _excelApp.SheetFollowHyperlinkEvent                 += (NetOffice.ICOMObject sh, Excel.Hyperlink target) => LogEvent("SheetFollowHyperlink");
            _excelApp.SheetLensGalleryRenderCompleteEvent       += (NetOffice.ICOMObject sh) => LogEvent("SheetLensGalleryRenderComplete");
            _excelApp.SheetSelectionChangeEvent                 += (NetOffice.ICOMObject sh, Excel.Range target) => LogEvent("SheetSelectionChange", TrySafeText(() => target.Address));
            _excelApp.WorkbookNewSheetEvent                     += (Excel.Workbook wb, NetOffice.ICOMObject sh) => LogEvent("WorkbookNewSheet", wb.Name);
            // XML
            _excelApp.WorkbookAfterXmlExportEvent               += (Excel.Workbook wb, Excel.XmlMap map, string url, Excel.Enums.XlXmlExportResult result) => LogEvent("WorkbookAfterXmlExport", wb.Name);
            _excelApp.WorkbookAfterXmlImportEvent               += (Excel.Workbook wb, Excel.XmlMap map, bool isRefresh, Excel.Enums.XlXmlImportResult result) => LogEvent("WorkbookAfterXmlImport", wb.Name);
            _excelApp.WorkbookBeforeXmlExportEvent              += (Excel.Workbook wb, Excel.XmlMap map, string url, ref bool c) => LogEvent("WorkbookBeforeXmlExport", wb.Name);
            _excelApp.WorkbookBeforeXmlImportEvent              += (Excel.Workbook wb, Excel.XmlMap map, string url, bool isRefresh, ref bool c) => LogEvent("WorkbookBeforeXmlImport", wb.Name);
            // Window
            _excelApp.WindowActivateEvent                       += (Excel.Workbook wb, Excel.Window wn) => { LogEvent("WindowActivate", wb.Name); BeginInvoke(() => RefreshDocs()); };
            _excelApp.WindowDeactivateEvent                     += (Excel.Workbook wb, Excel.Window wn) => LogEvent("WindowDeactivate", wb.Name);
            _excelApp.WindowResizeEvent                         += (Excel.Workbook wb, Excel.Window wn) => LogEvent("WindowResize", wb.Name);
            // Protected View
            _excelApp.ProtectedViewWindowOpenEvent              += (Excel.ProtectedViewWindow pvw) => LogEvent("ProtectedViewWindowOpen");
            _excelApp.ProtectedViewWindowBeforeEditEvent        += (Excel.ProtectedViewWindow pvw, ref bool c) => LogEvent("ProtectedViewWindowBeforeEdit");
            _excelApp.ProtectedViewWindowBeforeCloseEvent       += (Excel.ProtectedViewWindow pvw, Excel.Enums.XlProtectedViewCloseReason r, ref bool c) => LogEvent("ProtectedViewWindowBeforeClose");
            _excelApp.ProtectedViewWindowResizeEvent            += (Excel.ProtectedViewWindow pvw) => LogEvent("ProtectedViewWindowResize");
            _excelApp.ProtectedViewWindowActivateEvent          += (Excel.ProtectedViewWindow pvw) => LogEvent("ProtectedViewWindowActivate");
            _excelApp.ProtectedViewWindowDeactivateEvent        += (Excel.ProtectedViewWindow pvw) => LogEvent("ProtectedViewWindowDeactivate");
        }

        private void UnwireEvents()
        {
            // NetOffice unsubscribes all events automatically when the COM wrapper is disposed.
        }

        private void RefreshDocs()
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
    }
}
