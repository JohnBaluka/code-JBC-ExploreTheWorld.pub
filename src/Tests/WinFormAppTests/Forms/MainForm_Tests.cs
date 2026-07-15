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

        MainWindow!.Title.Should().Contain("ETW WinFormApp");
    }

    [Fact]
    public void MainWindow_ContainsBlazorWebViewPanel()
    {
        LaunchApp();

        // The single ExploreTheWorld_Form hosts a BlazorWebView (WebView2) fill-docked panel.
        var webViewPane = MainWindow!.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Pane));

        webViewPane.Should().NotBeNull();
    }

    [Fact]
    public void MainWindow_WebView_RendersDocumentContent()
    {
        LaunchApp();

        // WebView2 exposes the rendered page as a UIA Document once Blazor has painted —
        // this is the check that would have caught the blank-screenshot problem.
        var document = Retry.WhileNull(() =>
                MainWindow!.FindFirstDescendant(cf =>
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document)),
            TimeSpan.FromSeconds(15)).Result;

        document.Should().NotBeNull("the BlazorWebView should expose its rendered page as a UIA Document");
    }
}
