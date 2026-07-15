using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.AccessDbTests.Database;

[Trait("Category", "AccessDb")]
public class AccessApp_Tests : FlaUIAccessTestBase
{
    public AccessApp_Tests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void Access_Launches_OpensDatabase()
    {
        if (!IsAccessAvailable) return;

        LaunchAccess();

        MainWindow.Should().NotBeNull();
        MainWindow!.IsAvailable.Should().BeTrue();
        // The window title is the database's Application Title — the branding display name
        // "Explore the World" (three words), not the one-word "ExploreTheWorld" file name.
        MainWindow.Title.Should().Contain("Explore the World");
    }

    // The Navigation Pane shows the custom "Explore The World" category, so entries
    // carry the forms' display captions rather than their object names.

    [Fact]
    public void NavigationPane_ContainsCountriesNowForm()
    {
        if (!IsAccessAvailable) return;

        LaunchAccess();

        var navItem = WaitForElement("CountriesNow Data");

        navItem.Should().NotBeNull("the Navigation Pane should list the CountriesNow Data form");
    }

    [Fact]
    public void NavigationPane_ContainsMsOfficeForms()
    {
        if (!IsAccessAvailable) return;

        LaunchAccess();

        WaitForElement("Watcher - Word").Should().NotBeNull("the Navigation Pane should list the Word watcher form");
        WaitForElement("Watcher - Excel").Should().NotBeNull("the Navigation Pane should list the Excel watcher form");
        WaitForElement("Watcher - PowerPoint").Should().NotBeNull("the Navigation Pane should list the PowerPoint watcher form");
    }
}
