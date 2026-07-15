using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.UnitTests.BL;

public class MsOfficeExportManager_Tests
{
    private static readonly System.Collections.Generic.List<MsOfficeCountry_Row> Countries = new()
    {
        new MsOfficeCountry_Row("Australia", "AU", "AUS")
    };

    private readonly Mock<MsOfficeExportRepoFactory__Interface> _mockFactory;
    private readonly Mock<MsWord__Repo__Interface> _mockWordRepo;
    private readonly Mock<MsExcel__Repo__Interface> _mockExcelRepo;
    private readonly Mock<MsPowerPoint__Repo__Interface> _mockPowerPointRepo;
    private readonly MsOfficeExportManager__Service _sut;

    public MsOfficeExportManager_Tests()
    {
        _mockFactory        = new Mock<MsOfficeExportRepoFactory__Interface>();
        _mockWordRepo       = new Mock<MsWord__Repo__Interface>();
        _mockExcelRepo      = new Mock<MsExcel__Repo__Interface>();
        _mockPowerPointRepo = new Mock<MsPowerPoint__Repo__Interface>();

        _mockFactory.Setup(f => f.GetWordRepo(It.IsAny<string>())).Returns(_mockWordRepo.Object);
        _mockFactory.Setup(f => f.GetExcelRepo(It.IsAny<string>())).Returns(_mockExcelRepo.Object);
        _mockFactory.Setup(f => f.GetPowerPointRepo(It.IsAny<string>())).Returns(_mockPowerPointRepo.Object);

        _sut = new MsOfficeExportManager__Service(_mockFactory.Object);
    }

    // ── ExportAsync ───────────────────────────────────────────────────────────────

    [Theory]
    [InlineData("Word")]
    [InlineData("Excel")]
    [InlineData("PowerPoint")]
    public async Task ExportAsync_RoutesToRepoForExportType(string exportType)
    {
        await _sut.ExportAsync(Countries, exportType, "OpenXML", @"C:\out\file", _ => { });

        _mockFactory.Verify(f => f.GetWordRepo("OpenXML"),       exportType == "Word"       ? Times.Once() : Times.Never());
        _mockFactory.Verify(f => f.GetExcelRepo("OpenXML"),      exportType == "Excel"      ? Times.Once() : Times.Never());
        _mockFactory.Verify(f => f.GetPowerPointRepo("OpenXML"), exportType == "PowerPoint" ? Times.Once() : Times.Never());
    }

    [Fact]
    public async Task ExportAsync_Word_DelegatesToWordRepoExport()
    {
        await _sut.ExportAsync(Countries, "Word", "OpenXML", @"C:\out\file.docx", _ => { });

        _mockWordRepo.Verify(r => r.ExportAsync(Countries, @"C:\out\file.docx", It.IsAny<Action<string>>()), Times.Once);
    }

    [Fact]
    public async Task ExportAsync_UnsupportedExportType_Throws()
    {
        var act = () => _sut.ExportAsync(Countries, "Visio", "OpenXML", @"C:\out\file", _ => { });

        await act.Should().ThrowAsync<NotSupportedException>();
    }

    // ── WriteDocumentJsonAsync ────────────────────────────────────────────────────

    [Fact]
    public async Task WriteDocumentJsonAsync_Excel_DelegatesToExcelRepo()
    {
        await _sut.WriteDocumentJsonAsync("Excel", "OpenXML", @"C:\in.xlsx", @"C:\out.json", _ => { });

        _mockExcelRepo.Verify(r => r.WriteDocumentJsonAsync(@"C:\in.xlsx", @"C:\out.json", It.IsAny<Action<string>>()), Times.Once);
    }

    [Fact]
    public async Task WriteDocumentJsonAsync_UnsupportedExportType_Throws()
    {
        var act = () => _sut.WriteDocumentJsonAsync("Visio", "OpenXML", @"C:\in", @"C:\out.json", _ => { });

        await act.Should().ThrowAsync<NotSupportedException>();
    }

    // ── ExportToRunningAppAsync ───────────────────────────────────────────────────

    [Fact]
    public async Task ExportToRunningAppAsync_RepoWithoutRunningAppSupport_Throws()
    {
        // The plain mocks do not implement MsOfficeRunningAppExport__Repo__Interface,
        // so running-app export must be rejected.
        var act = () => _sut.ExportToRunningAppAsync("Word", new object(), Countries, _ => { });

        await act.Should().ThrowAsync<NotSupportedException>();
    }

    [Fact]
    public async Task ExportToRunningAppAsync_ResolvesNetOfficeRepo_AndDelegates()
    {
        var runningAppRepo = new Mock<MsWord__Repo__Interface>();
        var runningAppExport = runningAppRepo.As<MsOfficeRunningAppExport__Repo__Interface>();
        _mockFactory.Setup(f => f.GetWordRepo("NetOffice")).Returns(runningAppRepo.Object);

        var hostApplication = new object();
        await _sut.ExportToRunningAppAsync("Word", hostApplication, Countries, _ => { });

        runningAppExport.Verify(r => r.ExportToRunningAppAsync(hostApplication, Countries, It.IsAny<Action<string>>()), Times.Once);
    }

    [Fact]
    public async Task ExportToRunningAppAsync_UnsupportedExportType_Throws()
    {
        var act = () => _sut.ExportToRunningAppAsync("Visio", new object(), Countries, _ => { });

        await act.Should().ThrowAsync<NotSupportedException>();
    }

    // ── Flag image enrichment ─────────────────────────────────────────────────────

    [Fact]
    public async Task ExportAsync_WithFlagImageManager_EnrichesRowsWithFlagPng()
    {
        var flagBytes = new byte[] { 1, 2, 3 };

        // The flag manager is now a concrete service; drive it via its DL store interface —
        // a cache hit for "AU" makes GetFlagImagesAsync return the flag bytes.
        var mockStore = new Mock<FlagImageStore__Repo__Interface>();
        mockStore.Setup(s => s.GetImageAsync("AU")).ReturnsAsync(flagBytes);
        mockStore.Setup(s => s.GetImageFilePathAsync(It.IsAny<string>())).ReturnsAsync((string?)null);

        var flagManager = new FlagImageManager__Service(
            mockStore.Object,
            new Mock<FlagImageDownload__Repo__Interface>().Object,
            new CountriesNowSpaceManager__Service(
                new Mock<CountriesNowSpaceApi_Interface>().Object,
                new Mock<CountriesNowSpaceApiManager__Repo__Interface>().Object));

        var sut = new MsOfficeExportManager__Service(_mockFactory.Object, flagManager);

        System.Collections.Generic.IList<MsOfficeCountry_Row>? exported = null;
        _mockWordRepo
            .Setup(r => r.ExportAsync(It.IsAny<System.Collections.Generic.IList<MsOfficeCountry_Row>>(),
                                      It.IsAny<string>(), It.IsAny<Action<string>>()))
            .Callback<System.Collections.Generic.IList<MsOfficeCountry_Row>, string, Action<string>>(
                (rows, _, _) => exported = rows)
            .Returns(Task.CompletedTask);

        await sut.ExportAsync(Countries, "Word", "OpenXML", @"C:\out\file.docx", _ => { });

        exported.Should().NotBeNull();
        exported!.Should().ContainSingle(r => r.Iso2 == "AU").Which.FlagPng.Should().BeEquivalentTo(flagBytes);
    }

    [Fact]
    public async Task ExportAsync_WithoutFlagImageManager_PassesRowsThroughUnchanged()
    {
        await _sut.ExportAsync(Countries, "Word", "OpenXML", @"C:\out\file.docx", _ => { });

        // Same list instance — no enrichment copy is made when no manager is configured.
        _mockWordRepo.Verify(r => r.ExportAsync(Countries, @"C:\out\file.docx", It.IsAny<Action<string>>()), Times.Once);
    }
}
