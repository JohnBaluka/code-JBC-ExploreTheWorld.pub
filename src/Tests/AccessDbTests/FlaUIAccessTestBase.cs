using System.IO;
using FlaUI.Core.Capturing;
using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.AccessDbTests;

/// <summary>
/// Base class for FlaUI-based Microsoft Access tests. Launches MSACCESS.EXE with the
/// repository's <c>VBA\Access\ExploreTheWorld.accdb</c> opened read-only (the database
/// is set to compact on close, which would otherwise rewrite the tracked file on every
/// test run) and cleans up on disposal.
/// </summary>
public abstract class FlaUIAccessTestBase : IDisposable
{
    private static readonly string MsAccessPath =
        Environment.GetEnvironmentVariable("ETW_MSACCESS_PATH")
        ?? FindMsAccessExecutable();

    private static readonly string DatabasePath =
        Environment.GetEnvironmentVariable("ETW_ACCESS_DB_PATH")
        ?? Path.GetFullPath(Path.Combine(
               AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\..\..\VBA\Access\ExploreTheWorld.accdb"));

    private readonly string _testFolder;

    protected Application?   App        { get; private set; }
    protected UIA3Automation Automation { get; } = new UIA3Automation();
    protected Window?        MainWindow { get; private set; }

    /// <summary>False when Access is not installed or the .accdb is missing — tests return early.</summary>
    protected static bool IsAccessAvailable =>
        File.Exists(MsAccessPath) && File.Exists(DatabasePath);

    protected FlaUIAccessTestBase(ITestOutputHelper output)
    {
        _testFolder = UITestSettings.GetTestFolder(output);
    }

    protected void LaunchAccess()
    {
        App = Application.Launch(MsAccessPath, $"\"{DatabasePath}\" /ro");
        App.WaitWhileMainHandleIsMissing(TimeSpan.FromSeconds(30));

        // Access briefly reports its splash screen as the process main window; the real
        // database window replaces it, leaving any previously acquired Window element stale
        // — so re-acquire until it appears. The window title is the database's Application
        // Title (the branding display name "Explore the World"), not the one-word file name
        // "ExploreTheWorld" (code identity); accept either so the test tracks the branding.
        var databaseName = Path.GetFileNameWithoutExtension(DatabasePath);
        const string displayTitle = "Explore the World";
        MainWindow = Retry.WhileNull(() =>
            {
                var window = App.GetMainWindow(Automation, TimeSpan.FromSeconds(5));
                return window is { IsAvailable: true }
                       && (window.Title.Contains(displayTitle) || window.Title.Contains(databaseName))
                    ? window
                    : null;
            },
            TimeSpan.FromSeconds(30),
            ignoreException: true).Result;

        MainWindow.Should().NotBeNull($"the Access window for '{displayTitle}' should open");
        ResizeWindow();
        WaitForUiToSettle();
        TakeScreenshot("before");
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

    /// <summary>
    /// Finds a descendant of the Access window by its UIA name, retrying until it
    /// appears (Access builds the Navigation Pane and form UIA trees lazily).
    /// </summary>
    protected AutomationElement? WaitForElement(string name, int timeoutSeconds = 10)
        => Retry.WhileNull(
               () => MainWindow?.FindFirstDescendant(cf => cf.ByName(name)),
               TimeSpan.FromSeconds(timeoutSeconds),
               ignoreException: true).Result;

    /// <summary>
    /// Opens a form by double-clicking its Navigation Pane entry and retries until
    /// <paramref name="openedSentinel"/> — an element that exists only once the form is
    /// open (a button caption, a control label) — appears in the window. The pane shows
    /// the custom "Explore The World" category, whose entries carry the forms' display
    /// captions, so the sentinel must not equal the Navigation Pane entry name itself.
    /// Returns false when the form never opens.
    /// </summary>
    protected bool TryOpenNavigationPaneForm(string itemName, string openedSentinel)
    {
        var opened = Retry.WhileFalse(() =>
            {
                if (MainWindow?.FindFirstDescendant(cf => cf.ByName(openedSentinel)) != null)
                    return true;

                var item = FindFormNavigationItem(itemName);
                if (item is null) return false;

                try { item.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView(); } catch { }
                item.DoubleClick();
                Thread.Sleep(1000);
                return MainWindow?.FindFirstDescendant(cf => cf.ByName(openedSentinel)) != null;
            },
            TimeSpan.FromSeconds(30),
            TimeSpan.FromSeconds(1),
            ignoreException: true).Result;

        WaitForUiToSettle();
        return opened;
    }

    // Group headers the form entries can live under: the custom category group first
    // ("Explore The World"), then the classic object-type group ("Forms").
    private static readonly string[] FormGroupNames = { "Explore The World", "Forms" };

    private AutomationElement? FindFormNavigationItem(string itemName)
    {
        var items = MainWindow?.FindAllDescendants(cf => cf.ByName(itemName));
        var formItem = items?.FirstOrDefault(i => FormGroupNames.Any(g => HasAncestorNamed(i, g)));
        if (formItem != null) return formItem;

        // A collapsed group has no items in the UIA tree — expand it and rescan.
        foreach (var groupName in FormGroupNames)
        {
            var groupHeader = MainWindow?.FindFirstDescendant(cf => cf.ByName(groupName));
            if (groupHeader is null) continue;

            try { groupHeader.Patterns.ExpandCollapse.PatternOrDefault?.Expand(); }
            catch { groupHeader.DoubleClick(); }
            Thread.Sleep(500);

            items = MainWindow?.FindAllDescendants(cf => cf.ByName(itemName));
            formItem = items?.FirstOrDefault(i => HasAncestorNamed(i, groupName));
            if (formItem != null) return formItem;
        }

        return items is { Length: > 0 } ? items[items.Length - 1] : null;
    }

    private bool HasAncestorNamed(AutomationElement element, string name)
    {
        try
        {
            var walker = Automation.TreeWalkerFactory.GetRawViewWalker();
            var current = walker.GetParent(element);
            for (var depth = 0; current != null && depth < 10; depth++)
            {
                if (current.Properties.Name.ValueOrDefault == name) return true;
                current = walker.GetParent(current);
            }
        }
        catch { }
        return false;
    }

    /// <summary>
    /// Brings the Access window to the foreground (screenshots are screen captures, so
    /// an occluded window captures whatever covers it) and waits for the document area
    /// to finish painting before anything is captured or asserted.
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
        try { App?.WaitWhileBusy(TimeSpan.FromSeconds(10)); } catch { }
        Thread.Sleep(UITestSettings.SettleDelayMs);
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
        try
        {
            App?.Close();
            if (App is { HasExited: false })
                App.Kill();
        }
        catch { }
        App?.Dispose();
        Automation.Dispose();
    }

    private static string FindMsAccessExecutable()
    {
        var candidates = new[]
        {
            @"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"
        };

        return candidates.FirstOrDefault(File.Exists) ?? "MSACCESS.EXE";
    }
}
