using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.MauiAppTests.Pages;

[Trait("Category", "MauiApp")]
public class MainPage_Tests : FlaUITestBase
{
    public MainPage_Tests(ITestOutputHelper output) : base(output) { }

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

        MainWindow!.Title.Should().Contain("ETW MauiApp.WinUI");
    }

    [Fact]
    public void MainWindow_ContainsBlazorWebView()
    {
        LaunchApp();

        var webView = MainWindow!.FindFirstDescendant(cf =>
            cf.ByClassName("Microsoft.UI.Xaml.Controls.WebView2"));

        webView.Should().NotBeNull();
    }
}
