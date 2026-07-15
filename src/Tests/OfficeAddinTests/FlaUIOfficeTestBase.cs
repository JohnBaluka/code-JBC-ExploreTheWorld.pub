using System.IO;
using System.Runtime.InteropServices;
using FlaUI.Core.Capturing;
using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OfficeAddinTests;

/// <summary>
/// Base class for FlaUI-based Office add-in tests. Attaches to a running Office
/// application's UI Automation tree and cleans up on disposal.
/// </summary>
public abstract class FlaUIOfficeTestBase : IDisposable
{
    private readonly string _testFolder;

    protected Application?    OfficeApp  { get; private set; }
    protected UIA3Automation  Automation { get; } = new UIA3Automation();
    protected Window?         OfficeWindow { get; private set; }

    protected FlaUIOfficeTestBase(ITestOutputHelper output)
    {
        _testFolder = UITestSettings.GetTestFolder(output);
    }

    protected void LaunchOfficeApp(string processName, string windowTitle)
    {
        OfficeApp = Application.Attach(processName);
        OfficeWindow = Automation.GetDesktop()
            .FindFirstDescendant(cf => cf.ByName(windowTitle))
            ?.AsWindow();
        ResizeWindow();
        WaitForUiToSettle();
        TakeScreenshot("before");
    }

    protected void AttachToOfficeApp(string processName)
    {
        OfficeApp = Application.Attach(processName);
        OfficeWindow = OfficeApp.GetMainWindow(Automation, TimeSpan.FromSeconds(15));
        ResizeWindow();
        WaitForUiToSettle();
        TakeScreenshot("before");
    }

    /// <summary>
    /// Attaches to a running Office process and returns false if the main window
    /// is not accessible (e.g. process exists but has no open document/UI).
    /// </summary>
    protected bool TryAttachToOfficeApp(string processName)
    {
        try
        {
            OfficeApp = Application.Attach(processName);
            OfficeWindow = OfficeApp.GetMainWindow(Automation, TimeSpan.FromSeconds(15));
            if (OfficeWindow is null) return false;
            ResizeWindow();
            WaitForUiToSettle();
            TakeScreenshot("before");
            return true;
        }
        catch (COMException)   { return false; }
        catch (Exception e) when (e is System.ComponentModel.Win32Exception) { return false; }
    }

    private void ResizeWindow()
    {
        if (OfficeWindow is null) return;
        try
        {
            if (OfficeWindow.Patterns.Transform.IsSupported)
                OfficeWindow.Patterns.Transform.Pattern.Resize(UITestSettings.WindowWidth, UITestSettings.WindowHeight);
        }
        catch { }
    }

    /// <summary>
    /// Brings the Office window to the foreground (screenshots are screen captures, so
    /// an occluded window captures whatever covers it) and waits for the window to
    /// repaint before anything is captured or asserted.
    /// </summary>
    protected void WaitForUiToSettle()
    {
        if (OfficeWindow is null) return;
        try
        {
            OfficeWindow.SetForeground();
            OfficeWindow.Focus();
        }
        catch { }
        System.Threading.Thread.Sleep(UITestSettings.SettleDelayMs);
    }

    protected void TakeScreenshot(string name)
    {
        if (!UITestSettings.ScreenshotsEnabled || OfficeWindow is null) return;
        try
        {
            Directory.CreateDirectory(_testFolder);
            using var image = Capture.Element(OfficeWindow);
            image.ToFile(Path.Combine(_testFolder, $"{name}.png"));
        }
        catch { }
    }

    protected static bool IsOfficeRunning(string processName) =>
        System.Diagnostics.Process.GetProcessesByName(processName).Length > 0;

    public void Dispose()
    {
        TakeScreenshot("after");
        Automation.Dispose();
    }
}
