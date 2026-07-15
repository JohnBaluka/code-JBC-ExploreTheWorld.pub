namespace JBC.ExploreTheWorld.UnitTests.CL;

public class FlagImageUrl_Helper_Tests
{
    [Fact]
    public void GetPngThumbnailUrl_CommonsSvgUrl_ReturnsThumbnailPngUrl()
    {
        var result = FlagImageUrl_Helper.GetPngThumbnailUrl(
            "https://upload.wikimedia.org/wikipedia/commons/d/d4/Flag_of_Israel.svg");

        result.Should().Be(
            "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Flag_of_Israel.svg/330px-Flag_of_Israel.svg.png");
    }

    [Fact]
    public void GetPngThumbnailUrl_CustomWidth_UsesRequestedWidth()
    {
        var result = FlagImageUrl_Helper.GetPngThumbnailUrl(
            "https://upload.wikimedia.org/wikipedia/commons/b/ba/Flag_of_Germany.svg", 160);

        result.Should().Be(
            "https://upload.wikimedia.org/wikipedia/commons/thumb/b/ba/Flag_of_Germany.svg/160px-Flag_of_Germany.svg.png");
    }

    [Fact]
    public void GetPngThumbnailUrl_AlreadyThumbnailUrl_AppendsSizedPngRendition()
    {
        var result = FlagImageUrl_Helper.GetPngThumbnailUrl(
            "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Flag_of_Israel.svg");

        result.Should().Be(
            "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Flag_of_Israel.svg/330px-Flag_of_Israel.svg.png");
    }

    [Fact]
    public void GetPngThumbnailUrl_PngUrl_ReturnsUnchanged()
    {
        var result = FlagImageUrl_Helper.GetPngThumbnailUrl("https://example.com/flags/au.png");

        result.Should().Be("https://example.com/flags/au.png");
    }

    [Fact]
    public void GetPngThumbnailUrl_EnglishWikipediaSvgUrl_ReturnsThumbnailPngUrl()
    {
        // The CountriesNow data hosts several major countries' flags on English
        // Wikipedia (/wikipedia/en/) rather than Wikimedia Commons.
        var result = FlagImageUrl_Helper.GetPngThumbnailUrl(
            "https://upload.wikimedia.org/wikipedia/en/a/a4/Flag_of_the_United_States.svg");

        result.Should().Be(
            "https://upload.wikimedia.org/wikipedia/en/thumb/a/a4/Flag_of_the_United_States.svg/330px-Flag_of_the_United_States.svg.png");
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("https://example.com/flag.gif")]
    [InlineData("https://example.com/some/Flag_of_X.svg")] // SVG but not on a Wikimedia host
    public void GetPngThumbnailUrl_UnusableUrl_ReturnsNull(string? flagUrl)
    {
        FlagImageUrl_Helper.GetPngThumbnailUrl(flagUrl).Should().BeNull();
    }

    [Fact]
    public void GetPngThumbnailUrl_NonPositiveWidth_ReturnsNull()
    {
        FlagImageUrl_Helper.GetPngThumbnailUrl(
            "https://upload.wikimedia.org/wikipedia/commons/d/d4/Flag_of_Israel.svg", 0)
            .Should().BeNull();
    }
}
