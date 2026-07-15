using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Microsoft.Extensions.DependencyInjection;
using NetOffice;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Native;
using NetOffice.Tools.Native;
using JBC.ExploreTheWorld.AL.WinFormsLib;
using JBC.ExploreTheWorld.DL.MsOfficeApi_Impl;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.DL.MsJSInterop.Layout;
using JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs;
using Microsoft.AspNetCore.Components.WebView.WindowsForms;
using Radzen;
using Word = NetOffice.WordApi;

namespace JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn;

[ProgId("JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.Addin")]
[Guid("B1A2C3D4-E5F6-4789-ABCD-EF0123456789")]
[ComVisible(true)]
public class Addin : IDTExtensibility2, IRibbonExtensibility
{
    private static readonly string _addinOfficeRegistryKey = @"Software\Microsoft\Office\Word\Addins\";
    private static readonly string _progId               = "JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.Addin";
    private static readonly string _addinFriendlyName    = "Explore the World Word Add-in";
    private static readonly string _addinDescription     = "Explore the World Word Countries + Watcher";

    private Word.Application? _application;
    private ServiceProvider?  _services;
    private WatcherEvent_AppService? _watcherSvc;
    private static bool       _appDomainHandlerRegistered;

    // Each Blazor WebView form runs on its own dedicated STA thread with Application.Run().
    // This avoids the COM STA deadlock that occurs when WebView2 initialisation posts async
    // completion messages back to the Office STA message queue while that queue is blocked
    // inside the ribbon callback.
    private CountriesNowSpace_WebView_Form? _countriesNowBlazorForm;
    private MsWord_Watcher_WebView_Form?   _watcherBlazorForm;

    public Addin() { }

    // ------------------------------------------------------------------ //
    //  IDTExtensibility2
    // ------------------------------------------------------------------ //

    void IDTExtensibility2.OnConnection(object Application, NetOffice.Tools.ext_ConnectMode ConnectMode,
        object AddInInst, ref Array custom)
    {
        try
        {
            // comhost.dll does not set APP_CONTEXT_BASE_DIRECTORY; BlazorWebView calls
            // Path.GetRelativePath(AppContext.BaseDirectory, ...) inside StartWebViewCoreIfPossible()
            // which throws "The path is empty" when BaseDirectory is "". Fix it early.
            if (string.IsNullOrEmpty(AppContext.BaseDirectory))
            {
                var addinDir = Path.GetDirectoryName(typeof(Addin).Assembly.Location);
                if (!string.IsNullOrEmpty(addinDir))
                    AppContext.SetData("APP_CONTEXT_BASE_DIRECTORY",
                        addinDir + Path.DirectorySeparatorChar);
            }

            _application = new Word.Application(null, Application);
            LogInfo("OnConnection", $"thread={Environment.CurrentManagedThreadId}, apt={System.Threading.Thread.CurrentThread.GetApartmentState()}, BaseDir={AppContext.BaseDirectory}");

            var sc = new ServiceCollection();
            sc.AddTransient<CountriesNowSpaceApi_Interface, CountriesNowSpaceApi__Repo>();

            var dbProviderName = "SqliteDb";
            sc.AddSingleton(new JBC.ExploreTheWorld.AL.DbProvider_AppService { ProviderName = dbProviderName });
            switch (dbProviderName)
            {
                case "AccessDb":
                    sc.AddExploreTheWorldAccessDb(ExploreTheWorldAccessDb.DefaultDbPath);
                    break;
                default:
                    var dbPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "JBC", "ExploreTheWorld", "etw.db");
                    sc.AddExploreTheWorldSqliteDb($"Data Source={dbPath}");
                    break;
            }
            sc.AddTransient<CountriesNowSpaceManager__Service>();
            sc.AddWindowsFormsBlazorWebView();
            // Note: BlazorWebView.CreateFileProvider() uses only PhysicalFileProvider(wwwroot).
            // It does NOT read *.staticwebassets.runtime.json at runtime, and does NOT pick up
            // IFileProvider from DI. The _framework/ files (blazor.webview.js etc.) are copied
            // to wwwroot/_framework/ by the CopyBlazorWebViewFrameworkFiles build target.
            sc.AddRadzenComponents();
            sc.AddSingleton<WatcherEvent_AppService>();
            // Flag images — %LocalAppData% file cache with Wikimedia PNG download fallback.
            sc.AddSingleton<FlagImageStore__Repo__Interface, FlagImageStore_FileSystem__Repo>();
            sc.AddSingleton<FlagImageDownload__Repo__Interface, FlagImageDownload__Repo>();
            sc.AddTransient<FlagImageManager__Service>();
            sc.AddSingleton<MsOfficeExportRepoFactory__Interface, MsOfficeExportRepoFactory>();
            sc.AddSingleton<MsOfficeExportManager__Service>();
            sc.AddSingleton<JBC.ExploreTheWorld.AL.OfficeExport_AppService__Interface, OfficeExport_AppService>();
            sc.AddSingleton(new Layout_AppService { ShowSidebar = false });
            sc.AddSingleton<NewWindow_AppService__Interface, NullNewWindow_AppService>();
            // JS interop for the Blazor WebView forms (Main_Layout breakpoint badge, Country Slides)
            sc.AddScoped<Layout__Interop__Interface, Layout__Interop>();
            sc.AddScoped<RevealJs__Interop__Interface, RevealJs__Interop>();
            // Supply the platform-specific Office composition to the WinForms UI libraries.
            MsOfficeSaveAsJsonWriterProvider.Current  = new MsOfficeSaveAsJsonWriter();
            MsOfficeExportRepoFactoryProvider.Current = new MsOfficeExportRepoFactory();

            _services = sc.BuildServiceProvider();
            try { _services.EnsureExploreTheWorldDbCreated(); } catch { }

            _watcherSvc = _services.GetRequiredService<WatcherEvent_AppService>();
            _watcherSvc.IsOfficeAddinHost = true; // active document must stay open — hide OpenXML save
            WireWordEvents();

            RemoveLegacyTaskPaneProgIds();

            if (!_appDomainHandlerRegistered)
            {
                _appDomainHandlerRegistered = true;
                AppDomain.CurrentDomain.UnhandledException += (_, uea) =>
                {
                    var uex = uea.ExceptionObject is Exception exc ? exc
                        : new Exception(uea.ExceptionObject?.ToString() ?? "unknown");
                    LogError($"AppDomain.UnhandledException(IsTerminating={uea.IsTerminating})", uex);
                };
            }
        }
        catch (Exception ex) { LogError("OnConnection", ex); }
    }

    void IDTExtensibility2.OnDisconnection(NetOffice.Tools.ext_DisconnectMode RemoveMode, ref Array custom)
    {
        UnwireWordEvents();
        _watcherSvc?.Word_SetConnected(false);
        _watcherSvc = null;
        CloseFormSafe(_countriesNowBlazorForm); _countriesNowBlazorForm = null;
        CloseFormSafe(_watcherBlazorForm);      _watcherBlazorForm      = null;
        _services?.Dispose();                   _services               = null;
        _application?.Dispose();                _application            = null;
    }

    void IDTExtensibility2.OnStartupComplete(ref Array custom) { }
    void IDTExtensibility2.OnAddInsUpdate(ref Array custom) { }
    void IDTExtensibility2.OnBeginShutdown(ref Array custom) { }

    // ------------------------------------------------------------------ //
    //  IRibbonExtensibility
    // ------------------------------------------------------------------ //

    public string GetCustomUI(string RibbonID)
    {
        try
        {
            var asm = GetType().Assembly;
            string resourceName = $"{GetType().Namespace}.RibbonUI.xml";
            using Stream? stream = asm.GetManifestResourceStream(resourceName);
            if (stream is null) return string.Empty;
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch (Exception ex) { LogError("GetCustomUI", ex); return string.Empty; }
    }

    // ------------------------------------------------------------------ //
    //  Ribbon callbacks
    // ------------------------------------------------------------------ //

    public void CustomUI_OnLoad(NetOffice.OfficeApi.Native.IRibbonUI ribbonUI) { }

    public void OnOpenCountriesNowBlazorForm(IRibbonControl control)
    {
        try
        {
            if (_services == null) return;
            LogInfo("OnOpenCountriesNowBlazorForm", "invoked");

            if (_countriesNowBlazorForm is { IsDisposed: false } existing)
            {
                // Form is already running on its own thread; marshal Show/BringToFront to it.
                existing.BeginInvoke(new Action(() =>
                {
                    try { if (!existing.Visible) existing.Show(); else existing.BringToFront(); }
                    catch (Exception ex) { LogError("OnOpenCountriesNowBlazorForm(BeginInvoke)", ex); }
                }));
                return;
            }

            // Create a dedicated STA thread so WebView2 gets a proper Windows message pump
            // and the ribbon callback (COM invocation) can return immediately.
            var services = _services;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    LogInfo("OnOpenCountriesNowBlazorForm", $"thread entered id={Environment.CurrentManagedThreadId}, apt={System.Threading.Thread.CurrentThread.GetApartmentState()}");
                    try { System.Windows.Forms.Application.SetUnhandledExceptionMode(System.Windows.Forms.UnhandledExceptionMode.CatchException, threadScope: true); } catch { }
                    System.Windows.Forms.Application.ThreadException += (_, te) =>
                        LogError("OnOpenCountriesNowBlazorForm(ThreadException)", te.Exception);
                    LogInfo("OnOpenCountriesNowBlazorForm", "constructing CountriesNowSpace_WebView_Form");
                    var form = new CountriesNowSpace_WebView_Form(services);
                    form.Text = "ETW MsOfficeWordVstoAddIn";
                    LogInfo("OnOpenCountriesNowBlazorForm", "form constructed");
                    _countriesNowBlazorForm = form;
                    form.Load       += (_, _) => LogInfo("OnOpenCountriesNowBlazorForm", "form.Load");
                    form.Shown      += (_, _) => LogInfo("OnOpenCountriesNowBlazorForm", "form.Shown");
                    form.FormClosed += (_, fce) => LogInfo("OnOpenCountriesNowBlazorForm", $"form.FormClosed reason={fce.CloseReason}");
                    LogInfo("OnOpenCountriesNowBlazorForm", "calling Application.Run");
                    System.Windows.Forms.Application.Run(form);
                    LogInfo("OnOpenCountriesNowBlazorForm", "Application.Run returned");
                }
                catch (Exception ex) { LogError("OnOpenCountriesNowBlazorForm(thread)", ex); }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }
        catch (Exception ex) { LogError("OnOpenCountriesNowBlazorForm", ex); }
    }

    public void OnSaveAsJson(IRibbonControl control)
    {
        try
        {
            if (_application == null) return;
            var doc = _application.ActiveDocument;
            if (doc == null) { LogInfo("OnSaveAsJson", "no active document"); return; }
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    using var dlg = new System.Windows.Forms.SaveFileDialog
                    {
                        Title      = "Save Word Document as JSON",
                        Filter     = "JSON Files (*.json)|*.json",
                        DefaultExt = "json",
                        FileName   = JBC.ExploreTheWorld.CL.SaveAsJson_Helper.BuildDefaultPath(doc.FullName, "MsWord", "NetOffice")
                    };
                    if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
                    MsOfficeJsonWriter_Helper.WriteWordJson(doc, dlg.FileName,
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

    public void OnSaveAsJsonVba(IRibbonControl control)
    {
        try
        {
            if (_application == null) return;
            _application.Run("MSO_MsWord.WriteActiveDocument");
        }
        catch (Exception ex) { LogError("OnSaveAsJsonVba", ex); }
    }

    public void OnOpenWatcherBlazorForm(IRibbonControl control)
    {
        try
        {
            if (_services == null) return;
            LogInfo("OnOpenWatcherBlazorForm", "invoked");

            if (_watcherBlazorForm is { IsDisposed: false } existing)
            {
                existing.BeginInvoke(new Action(() =>
                {
                    try { if (!existing.Visible) existing.Show(); else existing.BringToFront(); }
                    catch (Exception ex) { LogError("OnOpenWatcherBlazorForm(BeginInvoke)", ex); }
                }));
                return;
            }

            var services = _services;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    LogInfo("OnOpenWatcherBlazorForm", $"thread entered id={Environment.CurrentManagedThreadId}, apt={System.Threading.Thread.CurrentThread.GetApartmentState()}");
                    try { System.Windows.Forms.Application.SetUnhandledExceptionMode(System.Windows.Forms.UnhandledExceptionMode.CatchException, threadScope: true); } catch { }
                    System.Windows.Forms.Application.ThreadException += (_, te) =>
                        LogError("OnOpenWatcherBlazorForm(ThreadException)", te.Exception);
                    LogInfo("OnOpenWatcherBlazorForm", "constructing MsWord_Watcher_WebView_Form");
                    var form = new MsWord_Watcher_WebView_Form(services);
                    form.Text = "ETW MsOfficeWordVstoAddIn";
                    LogInfo("OnOpenWatcherBlazorForm", "form constructed");
                    _watcherBlazorForm = form;
                    form.Load       += (_, _) => LogInfo("OnOpenWatcherBlazorForm", "form.Load");
                    form.Shown      += (_, _) => LogInfo("OnOpenWatcherBlazorForm", "form.Shown");
                    form.FormClosed += (_, fce) => LogInfo("OnOpenWatcherBlazorForm", $"form.FormClosed reason={fce.CloseReason}");
                    LogInfo("OnOpenWatcherBlazorForm", "calling Application.Run");
                    System.Windows.Forms.Application.Run(form);
                    LogInfo("OnOpenWatcherBlazorForm", "Application.Run returned");
                }
                catch (Exception ex) { LogError("OnOpenWatcherBlazorForm(thread)", ex); }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }
        catch (Exception ex) { LogError("OnOpenWatcherBlazorForm", ex); }
    }

    // ------------------------------------------------------------------ //
    //  Watcher event wiring
    // ------------------------------------------------------------------ //

    private void WireWordEvents()
    {
        if (_application == null || _watcherSvc == null) return;
        try
        {
            var activeDoc = TrySafeGet(() => _application.ActiveDocument?.Name);
            var openFiles = GetOpenDocNames();
            _watcherSvc.Word_SetConnected(true, activeDoc);
            _watcherSvc.Word_UpdateOpenFiles(openFiles, activeDoc);
            _watcherSvc.Word_AppendLog($"Connected to Word {_application.Version}");

            _application.DocumentOpenEvent          += OnWord_DocumentOpen;
            _application.NewDocumentEvent           += OnWord_NewDocument;
            _application.DocumentBeforeCloseEvent   += OnWord_DocumentBeforeClose;
            _application.DocumentBeforeSaveEvent    += OnWord_DocumentBeforeSave;
            _application.WindowActivateEvent        += OnWord_WindowActivate;
            _application.WindowDeactivateEvent      += OnWord_WindowDeactivate;
            _application.WindowSelectionChangeEvent += OnWord_WindowSelectionChange;
        }
        catch (Exception ex) { LogError("WireWordEvents", ex); }
    }

    private void UnwireWordEvents()
    {
        if (_application == null) return;
        try
        {
            _application.DocumentOpenEvent          -= OnWord_DocumentOpen;
            _application.NewDocumentEvent           -= OnWord_NewDocument;
            _application.DocumentBeforeCloseEvent   -= OnWord_DocumentBeforeClose;
            _application.DocumentBeforeSaveEvent    -= OnWord_DocumentBeforeSave;
            _application.WindowActivateEvent        -= OnWord_WindowActivate;
            _application.WindowDeactivateEvent      -= OnWord_WindowDeactivate;
            _application.WindowSelectionChangeEvent -= OnWord_WindowSelectionChange;
        }
        catch { }
    }

    private void OnWord_DocumentOpen(Word.Document doc)
    {
        if (_watcherSvc == null) return;
        _watcherSvc.Word_AppendLog($"DocumentOpen: {TrySafeGet(() => doc.Name)}");
        _watcherSvc.Word_UpdateOpenFiles(GetOpenDocNames(), TrySafeGet(() => doc.Name));
    }

    private void OnWord_NewDocument(Word.Document doc)
    {
        if (_watcherSvc == null) return;
        _watcherSvc.Word_AppendLog($"NewDocument: {TrySafeGet(() => doc.Name)}");
        _watcherSvc.Word_UpdateOpenFiles(GetOpenDocNames(), TrySafeGet(() => doc.Name));
    }

    private void OnWord_DocumentBeforeClose(Word.Document doc, ref bool cancel)
        => _watcherSvc?.Word_AppendLog($"DocumentBeforeClose: {TrySafeGet(() => doc.Name)}");

    private void OnWord_DocumentBeforeSave(Word.Document doc, ref bool saveAsUI, ref bool cancel)
        => _watcherSvc?.Word_AppendLog($"DocumentBeforeSave: {TrySafeGet(() => doc.Name)}");

    private void OnWord_WindowActivate(Word.Document doc, Word.Window wn)
    {
        if (_watcherSvc == null) return;
        var name = TrySafeGet(() => doc.Name);
        _watcherSvc.Word_AppendLog($"WindowActivate: {name}");
        _watcherSvc.Word_UpdateOpenFiles(GetOpenDocNames(), name);
    }

    private void OnWord_WindowDeactivate(Word.Document doc, Word.Window wn)
        => _watcherSvc?.Word_AppendLog($"WindowDeactivate: {TrySafeGet(() => doc.Name)}");

    private void OnWord_WindowSelectionChange(Word.Selection sel)
        => _watcherSvc?.Word_AppendLog("WindowSelectionChange");

    private List<string> GetOpenDocNames()
    {
        var names = new List<string>();
        try
        {
            foreach (Word.Document d in _application!.Documents)
                names.Add(TrySafeGet(() => d.Name) ?? "");
        }
        catch { }
        return names;
    }

    private static string? TrySafeGet(Func<string?> fn)
    {
        try { return fn(); } catch { return null; }
    }

    // ------------------------------------------------------------------ //
    //  Helpers
    // ------------------------------------------------------------------ //

    /// <summary>
    /// Thread-safe close: if the form is on another thread, BeginInvoke Close() to it.
    /// </summary>
    private static void CloseFormSafe(System.Windows.Forms.Form? form)
    {
        if (form == null || form.IsDisposed) return;
        try
        {
            if (form.InvokeRequired)
                form.BeginInvoke(new Action(() => { try { form.Close(); } catch { } }));
            else
                form.Close();
        }
        catch { }
    }

    // ------------------------------------------------------------------ //
    //  Legacy task-pane cleanup
    // ------------------------------------------------------------------ //

    private static void RemoveLegacyTaskPaneProgIds()
    {
        string[] legacyProgIds =
        [
            "JBC.ExploreTheWorld.AL.WinFormApp.CountriesNowSpace_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.CountriesNowSpace_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsWord_Watcher_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsWord_Watcher_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsExcel_Watcher_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsExcel_Watcher_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsPowerPoint_Watcher_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsPowerPoint_Watcher_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.CountriesNowSpace_UserControl",
            "JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.MsWord_Watcher_UserControl",
            "JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.CountriesNowSpace_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.MsWord_Watcher_WebView_UserControl",
        ];
        foreach (string progId in legacyProgIds)
        {
            try { Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\{progId}", throwOnMissingSubKey: false); }
            catch { }
        }
    }

    // ------------------------------------------------------------------ //
    //  COM registration
    // ------------------------------------------------------------------ //

    [ComRegisterFunction, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
    public static void RegisterFunction(Type type)
    {
        try
        {
            using var rk = Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _progId);
            rk.SetValue("LoadBehavior", 3);
            rk.SetValue("FriendlyName", _addinFriendlyName);
            rk.SetValue("Description", _addinDescription);
            using var lb = Registry.CurrentUser.CreateSubKey(
                @"Software\Classes\Interface\{000C0601-0000-0000-C000-000000000046}");
            if (lb.GetValue("") is null)
                lb.SetValue("", "Office .NET Framework Lockback Bypass Key");
        }
        catch (Exception ex) { LogError("RegisterFunction", ex); }
    }

    [ComUnregisterFunction, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
    public static void UnregisterFunction(Type type)
    {
        try { Registry.CurrentUser.DeleteSubKeyTree(_addinOfficeRegistryKey + _progId, throwOnMissingSubKey: false); }
        catch (Exception ex) { LogError("UnregisterFunction", ex); }
    }

    // ------------------------------------------------------------------ //
    //  Error logging
    // ------------------------------------------------------------------ //

    private static void LogError(string context, Exception ex)
    {
        try
        {
            string path = Path.Combine(Path.GetTempPath(), "JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.log");
            File.AppendAllText(path,
                $"[{DateTime.Now:u}] ERROR {context}{Environment.NewLine}{ex}{Environment.NewLine}---{Environment.NewLine}");
        }
        catch { }
    }

    private static void LogInfo(string context, string message)
    {
        try
        {
            string path = Path.Combine(Path.GetTempPath(), "JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.log");
            File.AppendAllText(path,
                $"[{DateTime.Now:u}] INFO {context}: {message}{Environment.NewLine}");
        }
        catch { }
    }
}
