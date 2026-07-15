using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OqtaneTests.Countries;

/// <summary>
/// Verifies the CountriesNow Oqtane module (CountriesNow__Module.Client/Server) renders
/// CountriesNow__Component with its full DI stack. A missing ClientStartup registration
/// does not just break the module — it crashes the whole Blazor Server circuit
/// (see docs/architecture.md, Oqtane Module Architecture).
/// </summary>
[Trait("Category", "Oqtane")]
[Collection("Oqtane")]
public class CountriesNow_Tests : PlaywrightTestBase
{
    public CountriesNow_Tests(ITestOutputHelper output, OqtaneServerFixture _) : base(output) { }

    [Fact]
    public async Task CountriesNow_Page_LoadsSuccessfully()
    {
        var response = await NavigateAsync($"{BaseUrl}/countries-now");

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }

    [Fact]
    public async Task CountriesNow_Module_RendersWithoutDiErrors()
    {
        await NavigateAsync($"{BaseUrl}/countries-now");

        // "Cannot provide a value for property ..." (missing ClientStartup registration)
        // crashes the circuit and surfaces as a console error, not an HTTP failure.
        ConsoleErrors.Should().NotContain(
            error => error.Contains("Cannot provide a value for property"),
            "every [Inject] dependency of CountriesNow__Component must be registered in ClientStartup");
    }

    [Fact]
    public async Task CountriesNow_Grid_AutoLoadsCountries()
    {
        await NavigateAsync($"{BaseUrl}/countries-now");

        var firstGrid = Page.Locator(".rz-data-grid").First;
        await Expect(firstGrid).ToBeVisibleAsync(new LocatorAssertionsToBeVisibleOptions { Timeout = 20_000 });

        // The countries grid populates from the DB (or API fallback) on load.
        var rows = Page.Locator(".rz-data-grid").First.Locator("tbody tr");
        int count = 0;
        for (var attempt = 0; attempt < 20 && count == 0; attempt++)
        {
            await Page.WaitForTimeoutAsync(500);
            count = await rows.CountAsync();
        }
        count.Should().BeGreaterThan(0, "the countries grid should auto-load rows");
    }

    [Fact]
    public async Task CountriesNow_RadzenDataGrid_JsInitializes()
    {
        await NavigateAsync($"{BaseUrl}/countries-now");

        // Radzen.Blazor.js must be loaded by the theme before the grid initializes;
        // otherwise RadzenDataGrid throws "Could not find 'Radzen.createDataGrid'".
        ConsoleErrors.Should().NotContain(
            error => error.Contains("Radzen") && error.Contains("undefined"),
            "Radzen.Blazor.js must be loaded before interactive Radzen components render");
    }
}
