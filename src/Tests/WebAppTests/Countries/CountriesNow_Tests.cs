using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.WebAppTests.Countries;

[Trait("Category", "WebApp")]
[Collection("WebApp")]
public class CountriesNow_Tests : PlaywrightTestBase
{
    public CountriesNow_Tests(ITestOutputHelper output, WebAppServerFixture _) : base(output) { }

    [Fact]
    public async Task CountriesNow_Page_LoadsSuccessfully()
    {
        await NavigateAsync($"{BaseUrl}/countries-now");

        var heading = Page.Locator("h1, h2, h3, h4, h5, h6").First;
        await Expect(heading).ToBeVisibleAsync();
    }

    [Fact]
    public async Task CountriesNow_Page_HasClearDatabaseButton()
    {
        await NavigateAsync($"{BaseUrl}/countries-now");

        var button = Page.GetByRole(AriaRole.Button, new PageGetByRoleOptions { Name = "Clear Database" });
        await Expect(button).ToBeVisibleAsync();
    }

    [Fact]
    public async Task CountriesNow_Page_AutoLoadsCountriesGrid()
    {
        await NavigateAsync($"{BaseUrl}/countries-now");

        var grid = Page.Locator(".rz-datatable, table").First;
        await Expect(grid).ToBeVisibleAsync(new LocatorAssertionsToBeVisibleOptions
        {
            Timeout = 15_000
        });
    }
}
