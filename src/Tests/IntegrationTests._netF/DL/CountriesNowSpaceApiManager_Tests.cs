using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;

namespace JBC.ExploreTheWorld.IntegrationTests.DL;

// Standalone _netF integration tests: IDbContextFactory<T> is available in EF Core 3.x but
// only provides CreateDbContext() (no async factory). TestDbContextFactory wraps DbContextOptions.
public class CountriesNowSpaceApiManager_Tests : IClassFixture<DatabaseFixture>
{
    private readonly DatabaseFixture _fixture;
    private readonly CountriesNowSpaceApiManager__Repo _sut;

    public CountriesNowSpaceApiManager_Tests(DatabaseFixture fixture)
    {
        _fixture = fixture;
        _sut = new CountriesNowSpaceApiManager__Repo(_fixture.CreateFactory());
    }

    // ── GetCountriesAsync ─────────────────────────────────────────────────────────

    [Fact]
    public async Task GetCountriesAsync_ReturnsSeededCountries()
    {
        var result = await _sut.GetCountriesAsync();

        result.Should().HaveCount(3);
        result.Select(c => c.Iso2).Should().Contain(new[] { "AU", "BR", "CA" });
    }

    // ── GetCountryByIso2Async ─────────────────────────────────────────────────────

    [Fact]
    public async Task GetCountryByIso2Async_KnownIso2_ReturnsCountry()
    {
        var result = await _sut.GetCountryByIso2Async("AU");

        result.Should().NotBeNull();
        result!.Country.Should().Be("Australia");
        result.Iso3.Should().Be("AUS");
    }

    [Fact]
    public async Task GetCountryByIso2Async_UnknownIso2_ReturnsNull()
    {
        var result = await _sut.GetCountryByIso2Async("ZZ");

        result.Should().BeNull();
    }

    // ── CreateCountryAsync ────────────────────────────────────────────────────────

    [Fact]
    public async Task CreateCountryAsync_PersistsNewCountry()
    {
        var repo = new CountriesNowSpaceApiManager__Repo(_fixture.CreateFactory());

        var newCountry = new cns_Country { Iso2 = "DE", Country = "Germany", Iso3 = "DEU" };
        await repo.CreateCountryAsync(newCountry);

        var retrieved = await repo.GetCountryByIso2Async("DE");
        retrieved.Should().NotBeNull();
        retrieved!.Country.Should().Be("Germany");

        await repo.DeleteCountryByIso2Async("DE");
    }

    // ── DeleteCountryByIso2Async ──────────────────────────────────────────────────

    [Fact]
    public async Task DeleteCountryByIso2Async_UnknownIso2_ReturnsFalse()
    {
        var result = await _sut.DeleteCountryByIso2Async("XX");

        result.Should().BeFalse();
    }
}
