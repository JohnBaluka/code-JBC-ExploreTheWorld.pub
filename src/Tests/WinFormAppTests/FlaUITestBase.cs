using System.IO;
using FlaUI.Core.Capturing;
using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.WinFormAppTests;

public abstract class FlaUITestBase : IDisposable
{
    private static readonly string AppPath =
        Environment.GetEnvironmentVariable("ETW_WINFORMAPP_PATH")
        ?? FindAppExecutable("JBC.ExploreTheWorld.AL.WinFormApp.exe");

    private readonly string _testFolder;

    protected Application?    App        { get; private set; }
    protected UIA3Automation  Automation { get; } = new UIA3Automation();
    protected Window?         MainWindow { get; private set; }

    protected FlaUITestBase(ITestOutputHelper output)
    {
        _testFolder = UITestSettings.GetTestFolder(output);
    }

    protected void LaunchApp()
    {
        App = Application.Launch(AppPath);
        MainWindow = App.GetMainWindow(Automation, TimeSpan.FromSeconds(10));
        ResizeWindow();
        WaitForUiToSettle();
        TakeScreenshot("before");
    }

    /// <summary>
    /// Brings the window to the foreground (screenshots are screen captures, so an
    /// occluded window captures whatever covers it) and waits for the BlazorWebView
    /// content to finish painting before anything is captured or asserted.
    /// </summary>
    protected void WaitForUiToSettle()
    {
        if (MainWindow is null) return;
        try
        {
            MainWindow.SetForeground();
            MainWindow.Focus();
        }
        catch { }
        try { App?.WaitWhileBusy(TimeSpan.FromSeconds(5)); } catch { }
        Thread.Sleep(UITestSettings.SettleDelayMs);
    }

    private void ResizeWindow()
    {
        if (MainWindow is null) return;
        try
        {
            if (MainWindow.Patterns.Transform.IsSupported)
                MainWindow.Patterns.Transform.Pattern.Resize(UITestSettings.WindowWidth, UITestSettings.WindowHeight);
        }
        catch { }
    }

    protected void TakeScreenshot(string name)
    {
        if (!UITestSettings.ScreenshotsEnabled || MainWindow is null) return;
        try
        {
            Directory.CreateDirectory(_testFolder);
            using var image = Capture.Element(MainWindow);
            image.ToFile(Path.Combine(_testFolder, $"{name}.png"));
        }
        catch { }
    }

    public void Dispose()
    {
        TakeScreenshot("after");
        App?.Close();
        App?.Dispose();
        Automation.Dispose();
    }

    private static string FindAppExecutable(string exeName)
    {
        var repoRoot = Path.GetFullPath(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\..\..\"));

        var candidates = new[]
        {
            Path.Combine(repoRoot, "src", "AL.WinFormApp", "bin", "Debug",   "net10.0-windows", exeName),
            Path.Combine(repoRoot, "src", "AL.WinFormApp", "bin", "Release", "net10.0-windows", exeName)
        };

        return candidates.FirstOrDefault(File.Exists) ?? exeName;
    }
}
