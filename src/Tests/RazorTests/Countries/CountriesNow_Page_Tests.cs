using JBC.ExploreTheWorld.AL.BlazorLib.Countries;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.RazorTests.Countries;

public class CountriesNow_Page_Tests : BlazorTestBase
{
    public CountriesNow_Page_Tests()
    {
        // The page injects the concrete CountriesNowSpaceManager__Service; build it from its
        // DL interfaces. A pre-seeded DB manager makes GetAllCountriesAsync return two
        // countries from the Database source without touching the API.
        var mockApi       = new Mock<CountriesNowSpaceApi_Interface>();
        var mockDbManager = new Mock<CountriesNowSpaceApiManager__Repo__Interface>();
        mockDbManager
            .Setup(db => db.GetCountriesAsync())
            .ReturnsAsync(new System.Collections.Generic.List<cns_Country>
            {
                new() { Country = "Australia", Iso2 = "AU", Iso3 = "AUS" },
                new() { Country = "Brazil",    Iso2 = "BR", Iso3 = "BRA" }
            });

        Services.AddScoped(_ =>
            new CountriesNowSpaceManager__Service(mockApi.Object, mockDbManager.Object));
    }

    [Fact]
    public void CountriesNow_Page_RendersWithoutThrowing()
    {
        var cut = Render<CountriesNow__Page>();

        cut.Markup.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void CountriesNow_Page_RendersLoadButton()
    {
        var cut = Render<CountriesNow__Page>();

        cut.Markup.Should().Contain("Load");
    }
}
