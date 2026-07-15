using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.WebAppTests;

public abstract class PlaywrightTestBase : IAsyncLifetime
{
    private readonly string _testFolder;
    private IBrowserContext _context = null!;

    protected IPlaywright PlaywrightInstance { get; private set; } = null!;
    protected IBrowser    Browser           { get; private set; } = null!;
    protected IPage       Page              { get; private set; } = null!;

    protected string BaseUrl { get; } =
        Environment.GetEnvironmentVariable("ETW_BASE_URL") ?? "https://localhost:7000";

    protected PlaywrightTestBase(ITestOutputHelper output)
    {
        _testFolder = UITestSettings.GetTestFolder(output);
    }

    public async Task InitializeAsync()
    {
        PlaywrightInstance = await Playwright.CreateAsync();
        Browser = await PlaywrightInstance.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Headless = UITestSettings.Headless
        });

        if (UITestSettings.ScreenshotsEnabled || UITestSettings.VideoEnabled)
            Directory.CreateDirectory(_testFolder);

        _context = await Browser.NewContextAsync(new BrowserNewContextOptions
        {
            IgnoreHTTPSErrors = true,
            ViewportSize      = new ViewportSize
            {
                Width  = UITestSettings.WindowWidth,
                Height = UITestSettings.WindowHeight
            },
            RecordVideoDir  = UITestSettings.VideoEnabled ? _testFolder : null,
            RecordVideoSize = UITestSettings.VideoEnabled
                ? new RecordVideoSize { Width = UITestSettings.WindowWidth, Height = UITestSettings.WindowHeight }
                : null
        });

        Page = await _context.NewPageAsync();
    }

    /// <summary>
    /// Navigates and waits for the network to go idle plus a settle delay so Blazor
    /// has painted, then captures the "before" screenshot. Use instead of
    /// <c>Page.GotoAsync</c> — a screenshot taken before navigation is always blank.
    /// </summary>
    protected async Task<IResponse?> NavigateAsync(string url)
    {
        var response = await Page.GotoAsync(url);

        try { await Page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 15_000 }); }
        catch (PlaywrightException) { }

        await Page.WaitForTimeoutAsync(UITestSettings.SettleDelayMs);

        if (UITestSettings.ScreenshotsEnabled)
            await TakeScreenshotAsync("before");

        return response;
    }

    public async Task DisposeAsync()
    {
        if (UITestSettings.ScreenshotsEnabled)
        {
            try { await TakeScreenshotAsync("after"); } catch { }
        }

        string? videoPath = null;
        if (UITestSettings.VideoEnabled && Page.Video != null)
        {
            try { videoPath = await Page.Video.PathAsync(); } catch { }
        }

        await _context.CloseAsync();

        if (videoPath != null && File.Exists(videoPath))
            File.Move(videoPath, Path.Combine(_testFolder, "recording.webm"), overwrite: true);

        await Browser.DisposeAsync();
        PlaywrightInstance.Dispose();
    }

    private async Task TakeScreenshotAsync(string name)
    {
        Directory.CreateDirectory(_testFolder);
        await Page.ScreenshotAsync(new PageScreenshotOptions
        {
            Path = Path.Combine(_testFolder, $"{name}.png")
        });
    }

    protected static ILocatorAssertions Expect(ILocator locator) => Assertions.Expect(locator);
    protected static IPageAssertions    Expect(IPage page)       => Assertions.Expect(page);
}
