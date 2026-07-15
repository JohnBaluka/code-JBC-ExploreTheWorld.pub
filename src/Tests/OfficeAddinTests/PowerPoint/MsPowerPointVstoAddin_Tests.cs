using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OfficeAddinTests.PowerPoint;

[Trait("Category", "OfficeAddin")]
[Trait("Host", "PowerPoint")]
public class MsPowerPointVstoAddin_Tests : FlaUIOfficeTestBase
{
    public MsPowerPointVstoAddin_Tests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void PowerPoint_IsRunning_AddinLoaded_RibbonTabVisible()
    {
        if (!IsOfficeRunning("POWERPNT"))
            return;

        if (!TryAttachToOfficeApp("POWERPNT"))
            return;

        var ribbonTab = OfficeWindow!.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.TabItem)
              .And(cf.ByName("ETW VSTO")));

        ribbonTab.Should().NotBeNull("the ETW VSTO ribbon tab should be loaded");
    }
}
