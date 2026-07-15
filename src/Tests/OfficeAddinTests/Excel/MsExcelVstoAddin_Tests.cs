using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OfficeAddinTests.Excel;

[Trait("Category", "OfficeAddin")]
[Trait("Host", "Excel")]
public class MsExcelVstoAddin_Tests : FlaUIOfficeTestBase
{
    public MsExcelVstoAddin_Tests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void Excel_IsRunning_AddinLoaded_RibbonTabVisible()
    {
        if (!IsOfficeRunning("EXCEL"))
            return;

        if (!TryAttachToOfficeApp("EXCEL"))
            return;

        var ribbonTab = OfficeWindow!.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.TabItem)
              .And(cf.ByName("ETW VSTO")));

        ribbonTab.Should().NotBeNull("the ETW VSTO ribbon tab should be loaded");
    }
}
