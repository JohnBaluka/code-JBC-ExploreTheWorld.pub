using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OqtaneTests.Theme;

/// <summary>
/// Verifies the ExploreTheWorld Radzen theme (src/AL.Oqtane/Theme):
/// RadzenLayout shell, left sidebar navigation, search/login in the header, and
/// Radzen stylesheet loading through Oqtane's resource pipeline (Oqtane has no
/// HeadOutlet, so &lt;HeadContent&gt; would silently not render).
/// </summary>
[Trait("Category", "Oqtane")]
[Collection("Oqtane")]
public class Theme_Tests : PlaywrightTestBase
{
    public Theme_Tests(ITestOutputHelper output, OqtaneServerFixture _) : base(output) { }

    [Fact]
    public async Task Home_Page_UsesRadzenLayout()
    {
        await NavigateAsync(BaseUrl);

        await Expect(Page.Locator(".rz-layout")).ToBeVisibleAsync();
        await Expect(Page.Locator(".rz-header")).ToBeVisibleAsync();
        await Expect(Page.Locator(".rz-body")).ToBeVisibleAsync();
    }

    [Fact]
    public async Task Sidebar_ShowsLeftNavigationMenu()
    {
        await NavigateAsync(BaseUrl);

        var menu = Page.Locator(".rz-sidebar .rz-panel-menu");
        await Expect(menu).ToBeVisibleAsync();
        await Expect(menu.GetByText("Countries Now")).ToBeVisibleAsync();
        await Expect(menu.GetByText("Country Slides")).ToBeVisibleAsync();
    }

    [Fact]
    public async Task Header_KeepsOqtaneSearchAndLoginControls()
    {
        await NavigateAsync(BaseUrl);

        var header = Page.Locator(".rz-header");
        // Oqtane's Search control renders a static-SSR form with a "keywords" input.
        await Expect(header.Locator("input[name='keywords']")).ToBeVisibleAsync();
        await Expect(header.GetByText("Login")).ToBeVisibleAsync();
    }

    [Fact]
    public async Task RadzenStylesheet_LoadsOnEveryThemedPage()
    {
        // Home has no module that references Radzen — the stylesheet must come from
        // the theme's Resources override, not from module resources.
        await NavigateAsync(BaseUrl);

        var radzenCssLinks = await Page.EvalOnSelectorAllAsync<string[]>(
            "link[rel=stylesheet]",
            "els => els.map(e => e.href).filter(h => h.includes('Radzen.Blazor/css'))");

        radzenCssLinks.Should().NotBeEmpty("the theme must load a Radzen stylesheet on pages without Radzen modules");
        radzenCssLinks.Should().OnlyContain(href => !href.Contains("-base.css"),
            "only the complete Radzen theme files may be used (repo UI-styling rules)");
    }

    [Fact]
    public async Task Header_Background_MatchesSidebar()
    {
        await NavigateAsync(BaseUrl);

        // theme.css redefines the header variables to the sidebar ones so the two stay
        // in sync in every selectable Radzen theme.
        var headerBackground = await Page.EvalOnSelectorAsync<string>(
            ".rz-header", "el => getComputedStyle(el).backgroundColor");
        var sidebarBackground = await Page.EvalOnSelectorAsync<string>(
            ".rz-sidebar", "el => getComputedStyle(el).backgroundColor");

        headerBackground.Should().Be(sidebarBackground);
    }

    [Fact]
    public async Task Sidebar_Toggle_CollapsesAndExpandsMenu()
    {
        await NavigateAsync(BaseUrl);

        var sidebar = Page.Locator(".rz-sidebar");
        await Expect(sidebar).ToHaveClassAsync(new System.Text.RegularExpressions.Regex(".*rz-sidebar-expanded.*"));

        // The theme renders statically, so the toggle is handled by the theme.js resource
        // (Blazor @onclick handlers never fire in Oqtane's Static render mode).
        await Page.Locator(".rz-sidebar-toggle").ClickAsync();
        await Expect(sidebar).ToHaveClassAsync(new System.Text.RegularExpressions.Regex(".*rz-sidebar-collapsed.*"),
            new LocatorAssertionsToHaveClassOptions { Timeout = 5_000 });

        await Page.Locator(".rz-sidebar-toggle").ClickAsync();
        await Expect(sidebar).ToHaveClassAsync(new System.Text.RegularExpressions.Regex(".*rz-sidebar-expanded.*"),
            new LocatorAssertionsToHaveClassOptions { Timeout = 5_000 });
    }
}
