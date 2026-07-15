using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.WebAppTests.Navigation;

[Trait("Category", "WebApp")]
[Collection("WebApp")]
public class Navigation_Tests : PlaywrightTestBase
{
    public Navigation_Tests(ITestOutputHelper output, WebAppServerFixture _) : base(output) { }

    [Fact]
    public async Task Home_Page_LoadsSuccessfully()
    {
        var response = await NavigateAsync(BaseUrl);

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }

    [Fact]
    public async Task CountrySlides_Page_LoadsSuccessfully()
    {
        await NavigateAsync($"{BaseUrl}/country-slides");

        var content = await Page.ContentAsync();
        content.Should().NotBeNullOrEmpty();
    }
}
