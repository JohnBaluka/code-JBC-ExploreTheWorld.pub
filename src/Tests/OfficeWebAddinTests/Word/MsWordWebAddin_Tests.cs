using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OfficeWebAddinTests.Word;

[Trait("Category", "OfficeWebAddin")]
[Trait("Host", "Word")]
[Collection("OfficeWebAddin")]
public class MsWordWebAddin_Tests : PlaywrightOfficeTestBase
{
    public MsWordWebAddin_Tests(ITestOutputHelper output, OfficeWebAddinServerFixture _) : base(output) { }

    [Fact]
    public async Task WordAddin_HomePage_LoadsSuccessfully()
    {
        await NavigateAsync(WordAddinBaseUrl);

        var content = await Page.ContentAsync();
        content.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public async Task WordAddin_HomePage_ContainsTaskPaneContent()
    {
        await NavigateAsync(WordAddinBaseUrl);

        var body = Page.Locator("body");
        await Expect(body).ToBeVisibleAsync();
        var html = await body.InnerHTMLAsync();
        html.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public async Task WordAddin_CountriesNowPage_LoadsSuccessfully()
    {
        var response = await NavigateAsync($"{WordAddinBaseUrl}/countries-now");

        response.Should().NotBeNull();
    }

    [Fact]
    public async Task WordAddin_ThemePage_LoadsSuccessfully()
    {
        var response = await NavigateAsync($"{WordAddinBaseUrl}/theme");

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }
}
