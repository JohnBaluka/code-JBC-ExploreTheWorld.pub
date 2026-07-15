using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OfficeWebAddinTests.PowerPoint;

[Trait("Category", "OfficeWebAddin")]
[Trait("Host", "PowerPoint")]
[Collection("OfficeWebAddin")]
public class MsPowerPointWebAddin_Tests : PlaywrightOfficeTestBase
{
    public MsPowerPointWebAddin_Tests(ITestOutputHelper output, OfficeWebAddinServerFixture _) : base(output) { }

    [Fact]
    public async Task PowerPointAddin_HomePage_LoadsSuccessfully()
    {
        var response = await NavigateAsync(PowerPointAddinBaseUrl);

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }

    [Fact]
    public async Task PowerPointAddin_CountriesNowPage_LoadsSuccessfully()
    {
        var response = await NavigateAsync($"{PowerPointAddinBaseUrl}/countries-now");

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }
}
