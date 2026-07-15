using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OfficeWebAddinTests.Excel;

[Trait("Category", "OfficeWebAddin")]
[Trait("Host", "Excel")]
[Collection("OfficeWebAddin")]
public class MsExcelWebAddin_Tests : PlaywrightOfficeTestBase
{
    public MsExcelWebAddin_Tests(ITestOutputHelper output, OfficeWebAddinServerFixture _) : base(output) { }

    [Fact]
    public async Task ExcelAddin_HomePage_LoadsSuccessfully()
    {
        var response = await NavigateAsync(ExcelAddinBaseUrl);

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }

    [Fact]
    public async Task ExcelAddin_CountriesNowPage_LoadsSuccessfully()
    {
        var response = await NavigateAsync($"{ExcelAddinBaseUrl}/countries-now");

        response.Should().NotBeNull();
        response!.Ok.Should().BeTrue();
    }
}
