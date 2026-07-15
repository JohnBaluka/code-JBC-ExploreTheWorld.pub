using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OfficeAddinTests.Word;

[Trait("Category", "OfficeAddin")]
[Trait("Host", "Word")]
public class MsWordVstoAddin_Tests : FlaUIOfficeTestBase
{
    public MsWordVstoAddin_Tests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void Word_IsRunning_AddinLoaded_RibbonTabVisible()
    {
        if (!IsOfficeRunning("WINWORD"))
            return;

        if (!TryAttachToOfficeApp("WINWORD"))
            return;

        var ribbonTab = OfficeWindow!.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.TabItem)
              .And(cf.ByName("ETW VSTO")));

        ribbonTab.Should().NotBeNull("the ETW VSTO ribbon tab should be loaded");
    }

    [Fact]
    public void Word_IsRunning_AddinLoaded_CanOpenCountriesForm()
    {
        if (!IsOfficeRunning("WINWORD"))
            return;

        if (!TryAttachToOfficeApp("WINWORD"))
            return;

        var ribbonTab = OfficeWindow!.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.TabItem)
              .And(cf.ByName("ETW VSTO")));

        ribbonTab.Should().NotBeNull();
        ribbonTab!.AsTabItem().Select();

        var countriesButton = OfficeWindow.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
              .And(cf.ByName("Countries Form (Blazor)")));

        countriesButton.Should().NotBeNull("Countries Form (Blazor) button should be on the ribbon");
    }
}
