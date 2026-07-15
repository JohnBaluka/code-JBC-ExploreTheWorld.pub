using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.AccessDbTests.Forms;

[Trait("Category", "AccessDb")]
public class CountriesNowDataForm_Tests : FlaUIAccessTestBase
{
    public CountriesNowDataForm_Tests(ITestOutputHelper output) : base(output) { }

    // The Navigation Pane entry carries the form's display caption ("CountriesNow Data"),
    // so the open-detection sentinel is the form's Load button — an element that only
    // exists once the form itself is open.

    [Fact]
    public void CountriesNowForm_Opens_FromNavigationPane()
    {
        if (!IsAccessAvailable) return;

        LaunchAccess();

        var opened = TryOpenNavigationPaneForm("CountriesNow Data", "Load");

        opened.Should().BeTrue("the CountriesNow Data form should open and show its Load button");
        TakeScreenshot("form-open");
    }

    [Fact]
    public void CountriesNowForm_ShowsLoadButton()
    {
        if (!IsAccessAvailable) return;

        LaunchAccess();

        if (!TryOpenNavigationPaneForm("CountriesNow Data", "Load")) return;

        var loadButton = WaitForElement("Load");

        loadButton.Should().NotBeNull("the form header should contain the Load button");
        TakeScreenshot("form-open");
    }

    [Fact]
    public void CountriesNowForm_ShowsExportControls()
    {
        if (!IsAccessAvailable) return;

        LaunchAccess();

        if (!TryOpenNavigationPaneForm("CountriesNow Data", "Load")) return;

        WaitForElement("Export").Should().NotBeNull("the form footer should contain the Export button");
        WaitForElement("Browse...").Should().NotBeNull("the form footer should contain the Browse... button");
        TakeScreenshot("form-open");
    }
}
