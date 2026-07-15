using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.WinFormAppTests.Forms;

[Trait("Category", "WinFormApp")]
public class MainForm_Tests : FlaUITestBase
{
    public MainForm_Tests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void App_Launches_ShowsMainWindow()
    {
        LaunchApp();

        MainWindow.Should().NotBeNull();
        MainWindow!.IsAvailable.Should().BeTrue();
    }

    [Fact]
    public void MainWindow_HasExpectedTitle()
    {
        LaunchApp();

        MainWindow!.Title.Should().Contain("ETW WinFormApp._netF");
    }

    [Fact]
    public void MainWindow_ContainsCountriesNowApiButton()
    {
        LaunchApp();

        var button = MainWindow!.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
              .And(cf.ByName("Countries Now API")));

        button.Should().NotBeNull();
    }

    [Fact]
    public void MainWindow_ContainsWatcherButton()
    {
        LaunchApp();

        var button = MainWindow!.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
              .And(cf.ByName("Watcher ▼")));

        button.Should().NotBeNull();
    }

    [Fact]
    public void CountriesNowApiButton_Click_OpensCountriesNowForm()
    {
        LaunchApp();

        var button = MainWindow!.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
              .And(cf.ByName("Countries Now API")));
        button.Should().NotBeNull();

        button!.AsButton().Invoke();
        WaitForUiToSettle();

        // Main_Form opens the modeless CountriesNowSpace_Form via Show().
        var countriesWindow = Retry.WhileNull(() =>
                App!.GetAllTopLevelWindows(Automation)
                    .FirstOrDefault(w => w.Title.Contains("Countries Now API")),
            TimeSpan.FromSeconds(10)).Result;

        countriesWindow.Should().NotBeNull("clicking 'Countries Now API' should open the CountriesNowSpace form");
        TakeScreenshot("countries-form-open");
    }
}
