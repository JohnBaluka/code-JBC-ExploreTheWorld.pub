using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.UnitTests.BL;

public class FlagImageManager_Tests
{
    private const string IsraelSvgUrl = "https://upload.wikimedia.org/wikipedia/commons/d/d4/Flag_of_Israel.svg";
    private const string IsraelPngUrl =
        "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Flag_of_Israel.svg/330px-Flag_of_Israel.svg.png";

    private static readonly byte[] PngBytes = { 1, 2, 3 };

    private readonly Mock<FlagImageStore__Repo__Interface> _mockStore;
    private readonly Mock<FlagImageDownload__Repo__Interface> _mockDownload;
    private readonly FlagImageManager__Service _sut;

    public FlagImageManager_Tests()
    {
        _mockStore    = new Mock<FlagImageStore__Repo__Interface>();
        _mockDownload = new Mock<FlagImageDownload__Repo__Interface>();

        _mockStore
            .Setup(s => s.GetImageFilePathAsync(It.IsAny<string>()))
            .ReturnsAsync((string?)null);

        // The CountriesNow collaborator is now the concrete service; feed it through its DL
        // interfaces. The DB manager returns a single cached flag, so GetCountryFlagsAsync
        // resolves the Israel SVG URL without hitting the API.
        var mockApi       = new Mock<CountriesNowSpaceApi_Interface>();
        var mockDbManager = new Mock<CountriesNowSpaceApiManager__Repo__Interface>();
        mockDbManager
            .Setup(db => db.GetCountryFlagsAsync())
            .ReturnsAsync(new System.Collections.Generic.List<cns_CountryFlag>
            {
                new() { Iso2 = "IL", Name = "Israel", Flag = IsraelSvgUrl }
            });
        var countriesNowManager =
            new CountriesNowSpaceManager__Service(mockApi.Object, mockDbManager.Object);

        _sut = new FlagImageManager__Service(
            _mockStore.Object, _mockDownload.Object, countriesNowManager);
    }

    // ── GetFlagImageAsync ─────────────────────────────────────────────────────────

    [Fact]
    public async Task GetFlagImageAsync_CacheHit_ReturnsDatabaseSource_WithoutDownloading()
    {
        _mockStore.Setup(s => s.GetImageAsync("IL")).ReturnsAsync(PngBytes);

        var result = await _sut.GetFlagImageAsync("il");

        result.Should().NotBeNull();
        result!.Iso2.Should().Be("IL");
        result.ImageBytes.Should().BeEquivalentTo(PngBytes);
        result.Source.Should().Be(DataSource_Enum.Database);
        _mockDownload.Verify(d => d.DownloadImageAsync(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public async Task GetFlagImageAsync_CacheMiss_DownloadsDerivedPngThumbnail_AndPersists()
    {
        _mockStore.Setup(s => s.GetImageAsync("IL")).ReturnsAsync((byte[]?)null);
        _mockDownload.Setup(d => d.DownloadImageAsync(IsraelPngUrl)).ReturnsAsync(PngBytes);

        var result = await _sut.GetFlagImageAsync("IL");

        result.Should().NotBeNull();
        result!.Source.Should().Be(DataSource_Enum.Api);
        result.ImageBytes.Should().BeEquivalentTo(PngBytes);
        _mockDownload.Verify(d => d.DownloadImageAsync(IsraelPngUrl), Times.Once);
        _mockStore.Verify(s => s.SaveImageAsync("IL", PngBytes), Times.Once);
    }

    [Fact]
    public async Task GetFlagImageAsync_UnknownIso2_ReturnsNull()
    {
        _mockStore.Setup(s => s.GetImageAsync(It.IsAny<string>())).ReturnsAsync((byte[]?)null);

        var result = await _sut.GetFlagImageAsync("XX");

        result.Should().BeNull();
        _mockDownload.Verify(d => d.DownloadImageAsync(It.IsAny<string>()), Times.Never);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    public async Task GetFlagImageAsync_EmptyIso2_ReturnsNull(string? iso2)
    {
        var result = await _sut.GetFlagImageAsync(iso2!);

        result.Should().BeNull();
    }

    // ── GetFlagImagesAsync ────────────────────────────────────────────────────────

    [Fact]
    public async Task GetFlagImagesAsync_SkipsFailedDownloads_AndReturnsResolvedImages()
    {
        _mockStore.Setup(s => s.GetImageAsync("IL")).ReturnsAsync(PngBytes);
        _mockStore.Setup(s => s.GetImageAsync("XX")).ThrowsAsync(new System.Net.Http.HttpRequestException("offline"));

        var result = await _sut.GetFlagImagesAsync(new[] { "IL", "XX" });

        result.Should().ContainKey("IL");
        result.Should().NotContainKey("XX");
    }

    // ── ClearCacheAsync ───────────────────────────────────────────────────────────

    [Fact]
    public async Task ClearCacheAsync_DelegatesToStore()
    {
        await _sut.ClearCacheAsync();

        _mockStore.Verify(s => s.ClearAsync(), Times.Once);
    }
}
