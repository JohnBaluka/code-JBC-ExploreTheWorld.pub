using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OqtaneTests.Navigation;

/// <summary>
/// Verifies client-side navigation between module pages. Oqtane Interactive render mode
/// runs the whole site on one Blazor Server circuit, so a crash on one module page kills
/// every other page in the session (regression guard for the DbProvider_AppService DI crash;
/// see docs/architecture.md, Oqtane Module Architecture).
/// </summary>
[Trait("Category", "Oqtane")]
[Collection("Oqtane")]
public class Navigation_Tests : PlaywrightTestBase
{
    public Navigation_Tests(ITestOutputHelper output, OqtaneServerFixture _) : base(output) { }

    [Fact]
    public async Task Home_Page_LoadsSuccessfully()
    {
        var response = await NavigateAsync(BaseUrl);

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }

    [Fact]
    public async Task Sidebar_Navigation_CountriesThenSlides_KeepsCircuitAlive()
    {
        await NavigateAsync(BaseUrl);

        var menu = Page.Locator(".rz-sidebar .rz-panel-menu");

        // Countries Now: the module must render (a DI failure here would crash the circuit).
        await menu.GetByText("Countries Now").ClickAsync();
        await Expect(Page.Locator(".rz-data-grid").First)
            .ToBeVisibleAsync(new LocatorAssertionsToBeVisibleOptions { Timeout = 20_000 });

        // Country Slides in the same circuit: must still render after visiting Countries Now.
        await menu.GetByText("Country Slides").ClickAsync();
        var slides = Page.Locator(".reveal .slides section");
        int count = 0;
        for (var attempt = 0; attempt < 30 && count == 0; attempt++)
        {
            await Page.WaitForTimeoutAsync(500);
            count = await slides.CountAsync();
        }
        count.Should().BeGreaterThan(2, "the slides module must render after navigating from the countries page");

        ConsoleErrors.Should().NotContain(
            error => error.Contains("Cannot provide a value for property") || error.Contains("ObjectDisposedException"),
            "no module may crash the shared Blazor Server circuit during navigation");
    }
}
