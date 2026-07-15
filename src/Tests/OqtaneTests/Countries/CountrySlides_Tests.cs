using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OqtaneTests.Countries;

/// <summary>
/// Verifies the CountrySlides Oqtane module (CountrySlides__Module.Client/Server):
/// reveal.js slideshow initialization and the explicit viewport height the module
/// wrapper must pass because Oqtane panes are auto-height
/// (see docs/architecture.md, CountrySlides__Module).
/// </summary>
[Trait("Category", "Oqtane")]
[Collection("Oqtane")]
public class CountrySlides_Tests : PlaywrightTestBase
{
    public CountrySlides_Tests(ITestOutputHelper output, OqtaneServerFixture _) : base(output) { }

    [Fact]
    public async Task CountrySlides_Page_LoadsSuccessfully()
    {
        var response = await NavigateAsync($"{BaseUrl}/country-slides");

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }

    [Fact]
    public async Task CountrySlides_RevealJs_RendersSlides()
    {
        await NavigateAsync($"{BaseUrl}/country-slides");

        var slides = Page.Locator(".reveal .slides section");
        int count = 0;
        for (var attempt = 0; attempt < 30 && count == 0; attempt++)
        {
            await Page.WaitForTimeoutAsync(500);
            count = await slides.CountAsync();
        }

        // Title slide + one per country + end slide.
        count.Should().BeGreaterThan(2, "reveal.js should render the title, country, and end slides");
    }

    [Fact]
    public async Task CountrySlides_Viewport_HasNonZeroHeight()
    {
        await NavigateAsync($"{BaseUrl}/country-slides");

        var wrapper = Page.Locator(".etw-slides-wrapper");
        await Expect(wrapper).ToBeVisibleAsync(new LocatorAssertionsToBeVisibleOptions { Timeout = 20_000 });

        // Oqtane panes are auto-height; without the explicit height Style on the module
        // wrapper the reveal.js viewport collapses to 0 px (regression guard).
        var box = await wrapper.BoundingBoxAsync();
        box.Should().NotBeNull();
        box!.Height.Should().BeGreaterThan(200, "the slides viewport must not collapse in the auto-height Oqtane pane");
    }

    [Fact]
    public async Task CountrySlides_Toolbar_ShowsCountryCount()
    {
        await NavigateAsync($"{BaseUrl}/country-slides");

        var badge = Page.Locator(".rz-badge", new PageLocatorOptions { HasTextRegex = new System.Text.RegularExpressions.Regex(@"\d+\s+COUNTRIES", System.Text.RegularExpressions.RegexOptions.IgnoreCase) });
        await Expect(badge.First).ToBeVisibleAsync(new LocatorAssertionsToBeVisibleOptions { Timeout = 20_000 });
    }
}
