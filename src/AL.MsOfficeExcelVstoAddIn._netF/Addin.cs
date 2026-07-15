using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using JBC.ExploreTheWorld.AL.WinFormsLib;
using JBC.ExploreTheWorld.DL.MsOfficeApi_Impl;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using NetOffice.Tools;
using NetOffice.ExcelApi.Tools;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace JBC.ExploreTheWorld.AL.MsOfficeExcelVstoAddIn;

[COMAddin("JBC ETW Excel Watcher", "ExploreTheWorld Excel event watcher", 3)]
[ProgId("JBC.ExploreTheWorld.AL.MsOfficeExcelVstoAddIn._netF.Addin")]
[Guid("F5E6F7A8-C9DA-4B23-EF01-234567890123")]
[CustomUI("RibbonUI.xml", true)]
// NOTE: no [CustomPane]. Excel is SDI — every workbook has its own top-level window and a Custom
// Task Pane belongs to exactly one window. We create panes per active window on demand (keyed by
// window Hwnd) instead of the single [CustomPane] pane. See "Per-window Custom Task Panes" below.
public class Addin : COMAddin
{
    // ------------------------------------------------------------------ //
    //  Per-window Custom Task Panes (Excel SDI)
    // ------------------------------------------------------------------ //
    private const string WatcherPaneProgId   = "JBC.ExploreTheWorld.AL.WinFormsLib.MsExcel_Watcher_UserControl";
    private const string CountriesPaneProgId = "JBC.ExploreTheWorld.AL.MsOfficeExcelVstoAddIn.CountriesNow_TaskPane_UserControl";
    private const string WatcherPaneTitle    = "ETW MsOfficeExcelVstoAddIn._netF — Watcher";
    private const string CountriesPaneTitle  = "ETW MsOfficeExcelVstoAddIn._netF — Countries API";

    private readonly Dictionary<int, Office.CustomTaskPane> _watcherPanes   = new Dictionary<int, Office.CustomTaskPane>();
    private readonly Dictionary<int, Office.CustomTaskPane> _countriesPanes = new Dictionary<int, Office.CustomTaskPane>();

    // When this add-in is COM-activated into EXCEL.EXE, dependencies are resolved against the
    // host's default AppDomain — whose probe path is Excel's install folder and the GAC, NOT this
    // add-in's bin directory. Some dependencies (e.g. Microsoft.Bcl.AsyncInterfaces, pulled in via
    // System.Text.Json) are also deployed at a higher version than callers reference, and a
    // .dll.config binding redirect is not honored for a COM-activated library. This resolver loads
    // any unresolved dependency from the add-in's own directory by simple name, ignoring the
    // requested version, which fixes both the probe-path and version-mismatch failures.
    static Addin()
    {
        AppDomain.CurrentDomain.AssemblyResolve += ResolveFromAddinDirectory;

        // Route NetOffice's internal diagnostics (which is where a failed ICTPFactory.CreateCTP
        // records its COM exception) to a log file so task-pane creation failures are visible.
        try
        {
            NetOffice.DebugConsole.Default.Mode     = NetOffice.DebugConsoleMode.LogFile;
            NetOffice.DebugConsole.Default.FileName = Path.Combine(Path.GetTempPath(),
                "JBC.ExploreTheWorld.AL.MsOfficeExcelVstoAddIn._netF.netoffice.log");
        }
        catch { }

        // Supply the platform-specific Office composition to the WinForms UI libraries,
        // which reference no DL repo _Impl project.
        MsOfficeSaveAsJsonWriterProvider.Current  = new MsOfficeSaveAsJsonWriter();
        MsOfficeExportRepoFactoryProvider.Current = new MsOfficeExportRepoFactory();
    }

    // The base implementation builds the ICTPFactory we need (exposed as TaskPaneFactory).
    // There are no [CustomPane]s to create — panes are made per workbook window on demand.
    public override void CTPFactoryAvailable(object CTPFactoryInst)
    {
        base.CTPFactoryAvailable(CTPFactoryInst);
        LogInfo("CTPFactoryAvailable", $"factory ready={TaskPaneFactory != null}");
    }

    // Ribbon load callback — base sets the RibbonUI property used to refresh toggle-button state.
    public override void CustomUI_OnLoad(NetOffice.OfficeApi.Native.IRibbonUI ribbonUI)
    {
        base.CustomUI_OnLoad(ribbonUI);
        LogInfo("CustomUI_OnLoad", "ribbon loaded");
    }

    // Surface NetOffice lifecycle errors (notably a failed CreateCTP) that the base COMAddin
    // otherwise only writes to its internal DebugConsole.
    protected override void OnError(NetOffice.Tools.ErrorMethodKind methodKind, Exception error)
        => LogError($"OnError({methodKind})", error);

    private static int SafeHwnd(Excel.Window window)
    {
        try { return window?.Hwnd ?? 0; } catch { return 0; }
    }

    private Excel.Window ActiveWindow()
    {
        try { return Application?.ActiveWindow; } catch { return null; }
    }

    // Get (or lazily create) the pane of the given kind for a specific workbook window.
    private Office.CustomTaskPane GetOrCreatePane(
        Dictionary<int, Office.CustomTaskPane> map, string progId, string title, int width, Excel.Window window)
    {
        int hwnd = SafeHwnd(window);
        if (map.TryGetValue(hwnd, out var existing) && existing != null)
        {
            try { _ = existing.Visible; return existing; }
            catch { map.Remove(hwnd); }
        }
        if (TaskPaneFactory == null)
        {
            LogError("GetOrCreatePane", new InvalidOperationException("TaskPaneFactory is null — CTPFactoryAvailable not received yet"));
            return null;
        }
        if (window == null) return null;
        try
        {
            var pane = TaskPaneFactory.CreateCTP(progId, title, window.UnderlyingObject) as Office.CustomTaskPane;
            if (pane == null)
            {
                LogError("GetOrCreatePane", new InvalidOperationException($"CreateCTP returned null for {progId}"));
                return null;
            }
            try { pane.DockPosition = Office.Enums.MsoCTPDockPosition.msoCTPDockPositionRight; } catch { }
            try { pane.Width = width; } catch { }
            pane.VisibleStateChangeEvent += OnPaneVisibleStateChanged;

            if (progId == WatcherPaneProgId)
            {
                try { if (pane.ContentControl is MsExcel_Watcher_UserControl uc) uc.InjectApplication(Application); }
                catch (Exception ex) { LogError("GetOrCreatePane(inject)", ex); }
            }

            map[hwnd] = pane;
            LogInfo("GetOrCreatePane", $"created {progId} for hwnd={hwnd}");
            return pane;
        }
        catch (Exception ex) { LogError("GetOrCreatePane", ex); return null; }
    }

    private void OnPaneVisibleStateChanged(Office._CustomTaskPane customTaskPaneInst)
    {
        try
        {
            if (RibbonUI == null) return;
            try { RibbonUI.InvalidateControl("etw_ExcelTaskPane"); } catch { }
            try { RibbonUI.InvalidateControl("etw_CountriesTaskPane_netF"); } catch { }
        }
        catch (Exception ex) { LogError("OnPaneVisibleStateChanged", ex); }
    }

    private static Assembly? ResolveFromAddinDirectory(object? sender, ResolveEventArgs args)
    {
        try
        {
            var requested = new AssemblyName(args.Name);
            // Resource assemblies and the satellite-resource probe are handled by the runtime.
            if (requested.Name == null || requested.Name.EndsWith(".resources", StringComparison.OrdinalIgnoreCase))
                return null;

            string addinDir = Path.GetDirectoryName(typeof(Addin).Assembly.Location) ?? string.Empty;
            string candidate = Path.Combine(addinDir, requested.Name + ".dll");
            if (!File.Exists(candidate)) return null;

            // Already loaded? Return the existing instance rather than loading a duplicate.
            foreach (var loaded in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (string.Equals(loaded.GetName().Name, requested.Name, StringComparison.OrdinalIgnoreCase))
                    return loaded;
            }

            return Assembly.LoadFrom(candidate);
        }
        catch (Exception ex)
        {
            LogError("ResolveFromAddinDirectory", ex);
            return null;
        }
    }

    // Called by onAction="OnOpenCountriesForm"
    public void OnOpenCountriesForm(Office.IRibbonControl control)
    {
        try
        {
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    var dbPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "JBC", "ExploreTheWorld", "etw.db");
                    var dbFactory = ExploreTheWorldSqliteDb.CreateFactory($"Data Source={dbPath}");
                    try { JBC.ExploreTheWorld.DL.CountriesNowSpaceData.ServiceCollectionExtensions.EnsureExploreTheWorldDbCreated(dbFactory); } catch { }
                    var api     = new CountriesNowSpaceApi__Repo();
                    var dbMgr   = new CountriesNowSpaceApiManager__Repo(dbFactory);
                    var manager = new CountriesNowSpaceManager__Service(api, dbMgr);
                    var form    = new CountriesNowSpace_Form(manager)
                    {
                        Text          = "Countries API",
                        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                    };
                    form.EnableAddinExportMode("Excel");
                    System.Windows.Forms.Application.Run(form);
                }
                catch (Exception ex) { LogError("OnOpenCountriesForm(thread)", ex); }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }
        catch (Exception ex) { LogError("OnOpenCountriesForm", ex); }
    }

    // Called by getPressed="OnGetPressedCountriesPane" — reports the pane state for the ACTIVE window.
    public bool OnGetPressedCountriesPane(Office.IRibbonControl control)
    {
        try
        {
            int hwnd = SafeHwnd(ActiveWindow());
            return _countriesPanes.TryGetValue(hwnd, out var pane) && pane != null && pane.Visible;
        }
        catch (Exception ex) { LogError("OnGetPressedCountriesPane", ex); return false; }
    }

    // Called by onAction="OnCheckCountriesPane" — show/hide the Countries pane on the ACTIVE window.
    public void OnCheckCountriesPane(Office.IRibbonControl control, bool pressed)
    {
        try
        {
            var win = ActiveWindow();
            if (win == null) { LogError("OnCheckCountriesPane", new InvalidOperationException("no active window")); return; }
            var pane = GetOrCreatePane(_countriesPanes, CountriesPaneProgId, CountriesPaneTitle, 480, win);
            if (pane == null) return;
            LogInfo("OnCheckCountriesPane", $"pressed={pressed} hwnd={SafeHwnd(win)}");
            pane.Visible = pressed;
        }
        catch (Exception ex) { LogError("OnCheckCountriesPane", ex); }
    }

    // Called by onAction="OnOpenWatcher"
    public void OnOpenWatcher(Office.IRibbonControl control)
    {
        try
        {
            var form = new System.Windows.Forms.Form
            {
                Text          = "Microsoft Excel - Watcher",
                Size          = new System.Drawing.Size(940, 660),
                MinimumSize   = new System.Drawing.Size(700, 480),
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            };
            var uc = new MsExcel_Watcher_UserControl { Dock = System.Windows.Forms.DockStyle.Fill };
            form.Controls.Add(uc);
            uc.InjectApplication(Application);
            form.Show();
        }
        catch (Exception ex) { LogError("OnOpenWatcher", ex); }
    }

    // Called by onAction="OnSaveAsJson"
    public void OnSaveAsJson(Office.IRibbonControl control)
    {
        try
        {
            var wb = Application?.ActiveWorkbook;
            if (wb == null) return;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    using var dlg = new System.Windows.Forms.SaveFileDialog
                    {
                        Title      = "Save Excel Workbook as JSON",
                        Filter     = "JSON Files (*.json)|*.json",
                        DefaultExt = "json",
                        FileName   = JBC.ExploreTheWorld.CL.SaveAsJson_Helper.BuildDefaultPath(wb.FullName, "MsExcel", "NetOffice")
                    };
                    if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
                    MsOfficeJsonWriter_Helper.WriteExcelJson(wb, dlg.FileName,
                        msg => LogInfo("OnSaveAsJson", msg));
                    System.Windows.Forms.MessageBox.Show(
                        $"Saved: {dlg.FileName}", "Save As JSON",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
                catch (Exception ex) { LogError("OnSaveAsJson(thread)", ex); }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }
        catch (Exception ex) { LogError("OnSaveAsJson", ex); }
    }

    // Called by onAction="OnSaveAsJsonVba"
    public void OnSaveAsJsonVba(Office.IRibbonControl control)
    {
        try { Application?.Run("MSO_MsExcel.WriteActiveWorkbook"); }
        catch (Exception ex) { LogError("OnSaveAsJsonVba", ex); }
    }

    // Called by getPressed="OnGetPressedPanelToggle" — reports the pane state for the ACTIVE window.
    public bool OnGetPressedPanelToggle(Office.IRibbonControl control)
    {
        try
        {
            int hwnd = SafeHwnd(ActiveWindow());
            return _watcherPanes.TryGetValue(hwnd, out var pane) && pane != null && pane.Visible;
        }
        catch (Exception ex) { LogError("OnGetPressedPanelToggle", ex); return false; }
    }

    // Called by onAction="OnCheckPanelToggle" — show/hide the Watcher pane on the ACTIVE window.
    public void OnCheckPanelToggle(Office.IRibbonControl control, bool pressed)
    {
        try
        {
            var win = ActiveWindow();
            if (win == null) { LogError("OnCheckPanelToggle", new InvalidOperationException("no active window")); return; }
            var pane = GetOrCreatePane(_watcherPanes, WatcherPaneProgId, WatcherPaneTitle, 400, win);
            if (pane == null) return;
            LogInfo("OnCheckPanelToggle", $"pressed={pressed} hwnd={SafeHwnd(win)}");
            pane.Visible = pressed;
        }
        catch (Exception ex) { LogError("OnCheckPanelToggle", ex); }
    }

    private static void LogError(string context, Exception ex)
    {
        try
        {
            string path = Path.Combine(Path.GetTempPath(), "JBC.ExploreTheWorld.AL.MsOfficeExcelVstoAddIn._netF.log");
            File.AppendAllText(path,
                $"[{DateTime.Now:u}] ERROR {context}{Environment.NewLine}{ex}{Environment.NewLine}---{Environment.NewLine}");
        }
        catch { }
    }

    private static void LogInfo(string context, string message)
    {
        try
        {
            string path = Path.Combine(Path.GetTempPath(), "JBC.ExploreTheWorld.AL.MsOfficeExcelVstoAddIn._netF.log");
            File.AppendAllText(path,
                $"[{DateTime.Now:u}] INFO {context}: {message}{Environment.NewLine}");
        }
        catch { }
    }
}
