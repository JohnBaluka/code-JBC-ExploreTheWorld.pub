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
using PowerPoint = NetOffice.PowerPointApi;

namespace JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn;

[ProgId("JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn.Addin")]
[Guid("D3C4E5F6-A7B8-4901-CDEF-012345678901")]
[ComVisible(true)]
public class Addin : IDTExtensibility2, IRibbonExtensibility
{
    private static readonly string _addinOfficeRegistryKey = @"Software\Microsoft\Office\PowerPoint\Addins\";
    private static readonly string _progId               = "JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn.Addin";
    private static readonly string _addinFriendlyName    = "Explore the World PowerPoint Add-in";
    private static readonly string _addinDescription     = "Explore the World PowerPoint Countries + Watcher";

    private PowerPoint.Application? _application;
    private ServiceProvider?        _services;
    private WatcherEvent_AppService?    _watcherSvc;
    private static bool             _appDomainHandlerRegistered;

    // Each Blazor WebView form runs on its own dedicated STA thread with Application.Run().
    private CountriesNowSpace_WebView_Form?       _countriesNowBlazorForm;
    private MsPowerPoint_Watcher_WebView_Form?    _watcherBlazorForm;

    public Addin() { }

    void IDTExtensibility2.OnConnection(object Application, NetOffice.Tools.ext_ConnectMode ConnectMode,
        object AddInInst, ref Array custom)
    {
        try
        {
            _application = new PowerPoint.Application(null, Application);
            LogInfo("OnConnection", $"thread={Environment.CurrentManagedThreadId}, apt={System.Threading.Thread.CurrentThread.GetApartmentState()}");

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
            WirePowerPointEvents();

            RemoveLegacyTaskPaneProgIds();

            // BlazorWebView resolves HostPage relative to AppContext.BaseDirectory.
            // The COM host may leave it empty; fix it to the add-in DLL directory.
            if (string.IsNullOrEmpty(AppContext.BaseDirectory))
                AppContext.SetData("APP_CONTEXT_BASE_DIRECTORY",
                    Path.GetDirectoryName(typeof(Addin).Assembly.Location)!);

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
        UnwirePowerPointEvents();
        _watcherSvc?.PowerPoint_SetConnected(false);
        _watcherSvc = null;
        CloseFormSafe(_countriesNowBlazorForm); _countriesNowBlazorForm = null;
        CloseFormSafe(_watcherBlazorForm);      _watcherBlazorForm      = null;
        _services?.Dispose();                   _services               = null;
        _application?.Dispose();                _application            = null;
    }

    void IDTExtensibility2.OnStartupComplete(ref Array custom) { }
    void IDTExtensibility2.OnAddInsUpdate(ref Array custom) { }
    void IDTExtensibility2.OnBeginShutdown(ref Array custom) { }

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

    public void CustomUI_OnLoad(NetOffice.OfficeApi.Native.IRibbonUI ribbonUI) { }

    public void OnOpenCountriesNowBlazorForm(IRibbonControl control)
    {
        try
        {
            if (_services == null) return;
            LogInfo("OnOpenCountriesNowBlazorForm", "invoked");

            if (_countriesNowBlazorForm is { IsDisposed: false } existing)
            {
                existing.BeginInvoke(new Action(() =>
                {
                    try { if (!existing.Visible) existing.Show(); else existing.BringToFront(); }
                    catch (Exception ex) { LogError("OnOpenCountriesNowBlazorForm(BeginInvoke)", ex); }
                }));
                return;
            }

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
                    form.Text = "ETW MsOfficePowerPointVstoAddIn";
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
            var prs = _application.ActivePresentation;
            if (prs == null) { LogInfo("OnSaveAsJson", "no active presentation"); return; }
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    using var dlg = new System.Windows.Forms.SaveFileDialog
                    {
                        Title      = "Save PowerPoint Presentation as JSON",
                        Filter     = "JSON Files (*.json)|*.json",
                        DefaultExt = "json",
                        FileName   = JBC.ExploreTheWorld.CL.SaveAsJson_Helper.BuildDefaultPath(prs.FullName, "MsPowerPoint", "NetOffice")
                    };
                    if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
                    MsOfficeJsonWriter_Helper.WritePowerPointJson(prs, dlg.FileName,
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
            _application.Run("MSO_MsPowerPoint.WriteActivePresentation");
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
                    LogInfo("OnOpenWatcherBlazorForm", "constructing MsPowerPoint_Watcher_WebView_Form");
                    var form = new MsPowerPoint_Watcher_WebView_Form(services);
                    form.Text = "ETW MsOfficePowerPointVstoAddIn";
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

    private void WirePowerPointEvents()
    {
        if (_application == null || _watcherSvc == null) return;
        try
        {
            var activeFile = TrySafeGet(() => _application.ActivePresentation?.Name);
            var openFiles  = GetOpenPresentationNames();
            _watcherSvc.PowerPoint_SetConnected(true, activeFile);
            _watcherSvc.PowerPoint_UpdateOpenFiles(openFiles, activeFile);
            _watcherSvc.PowerPoint_AppendLog($"Connected to PowerPoint {_application.Version}");

            _application.PresentationOpenEvent         += OnPowerPoint_PresentationOpen;
            _application.NewPresentationEvent          += OnPowerPoint_NewPresentation;
            _application.PresentationBeforeCloseEvent  += OnPowerPoint_PresentationBeforeClose;
            _application.PresentationBeforeSaveEvent   += OnPowerPoint_PresentationBeforeSave;
            _application.WindowActivateEvent           += OnPowerPoint_WindowActivate;
            _application.WindowDeactivateEvent         += OnPowerPoint_WindowDeactivate;
            _application.WindowSelectionChangeEvent    += OnPowerPoint_WindowSelectionChange;
        }
        catch (Exception ex) { LogError("WirePowerPointEvents", ex); }
    }

    private void UnwirePowerPointEvents()
    {
        if (_application == null) return;
        try
        {
            _application.PresentationOpenEvent         -= OnPowerPoint_PresentationOpen;
            _application.NewPresentationEvent          -= OnPowerPoint_NewPresentation;
            _application.PresentationBeforeCloseEvent  -= OnPowerPoint_PresentationBeforeClose;
            _application.PresentationBeforeSaveEvent   -= OnPowerPoint_PresentationBeforeSave;
            _application.WindowActivateEvent           -= OnPowerPoint_WindowActivate;
            _application.WindowDeactivateEvent         -= OnPowerPoint_WindowDeactivate;
            _application.WindowSelectionChangeEvent    -= OnPowerPoint_WindowSelectionChange;
        }
        catch { }
    }

    private void OnPowerPoint_PresentationOpen(PowerPoint.Presentation prs)
    {
        if (_watcherSvc == null) return;
        _watcherSvc.PowerPoint_AppendLog($"PresentationOpen: {TrySafeGet(() => prs.Name)}");
        _watcherSvc.PowerPoint_UpdateOpenFiles(GetOpenPresentationNames(), TrySafeGet(() => prs.Name));
    }

    private void OnPowerPoint_NewPresentation(PowerPoint.Presentation prs)
    {
        if (_watcherSvc == null) return;
        _watcherSvc.PowerPoint_AppendLog($"NewPresentation: {TrySafeGet(() => prs.Name)}");
        _watcherSvc.PowerPoint_UpdateOpenFiles(GetOpenPresentationNames(), TrySafeGet(() => prs.Name));
    }

    private void OnPowerPoint_PresentationBeforeClose(PowerPoint.Presentation prs, ref bool cancel)
        => _watcherSvc?.PowerPoint_AppendLog($"PresentationBeforeClose: {TrySafeGet(() => prs.Name)}");

    private void OnPowerPoint_PresentationBeforeSave(PowerPoint.Presentation prs, ref bool cancel)
        => _watcherSvc?.PowerPoint_AppendLog($"PresentationBeforeSave: {TrySafeGet(() => prs.Name)}");

    private void OnPowerPoint_WindowActivate(PowerPoint.Presentation prs, PowerPoint.DocumentWindow wn)
    {
        if (_watcherSvc == null) return;
        var name = TrySafeGet(() => prs.Name);
        _watcherSvc.PowerPoint_AppendLog($"WindowActivate: {name}");
        _watcherSvc.PowerPoint_UpdateOpenFiles(GetOpenPresentationNames(), name);
    }

    private void OnPowerPoint_WindowDeactivate(PowerPoint.Presentation prs, PowerPoint.DocumentWindow wn)
        => _watcherSvc?.PowerPoint_AppendLog($"WindowDeactivate: {TrySafeGet(() => prs.Name)}");

    private void OnPowerPoint_WindowSelectionChange(PowerPoint.Selection sel)
        => _watcherSvc?.PowerPoint_AppendLog("WindowSelectionChange");

    private List<string> GetOpenPresentationNames()
    {
        var names = new List<string>();
        try
        {
            foreach (PowerPoint.Presentation prs in _application!.Presentations)
                names.Add(TrySafeGet(() => prs.Name) ?? "");
        }
        catch { }
        return names;
    }

    private static string? TrySafeGet(Func<string?> fn)
    {
        try { return fn(); } catch { return null; }
    }

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

    private static void RemoveLegacyTaskPaneProgIds()
    {
        string[] legacyProgIds =
        [
            "JBC.ExploreTheWorld.AL.WinFormApp.CountriesNowSpace_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.CountriesNowSpace_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsPowerPoint_Watcher_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsPowerPoint_Watcher_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsWord_Watcher_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsWord_Watcher_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsExcel_Watcher_UserControl",
            "JBC.ExploreTheWorld.AL.WinFormApp.MsExcel_Watcher_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn.CountriesNowSpace_UserControl",
            "JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn.MsPowerPoint_Watcher_UserControl",
            "JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn.CountriesNowSpace_WebView_UserControl",
            "JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn.MsPowerPoint_Watcher_WebView_UserControl",
        ];
        foreach (string progId in legacyProgIds)
        {
            try { Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\{progId}", throwOnMissingSubKey: false); }
            catch { }
        }
    }

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

    private static void LogError(string context, Exception ex)
    {
        try
        {
            string path = Path.Combine(Path.GetTempPath(), "JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn.log");
            File.AppendAllText(path,
                $"[{DateTime.Now:u}] ERROR {context}{Environment.NewLine}{ex}{Environment.NewLine}---{Environment.NewLine}");
        }
        catch { }
    }

    private static void LogInfo(string context, string message)
    {
        try
        {
            string path = Path.Combine(Path.GetTempPath(), "JBC.ExploreTheWorld.AL.MsOfficePowerPointVstoAddIn.log");
            File.AppendAllText(path,
                $"[{DateTime.Now:u}] INFO {context}: {message}{Environment.NewLine}");
        }
        catch { }
    }
}
