using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;

namespace JBC.ExploreTheWorld.SqliteDbTests.DL;

public class CountriesNowSpaceApiManager_Tests : IClassFixture<SqliteDatabaseFixture>
{
    private readonly SqliteDatabaseFixture _fixture;
    private readonly CountriesNowSpaceApiManager__Repo _sut;

    public CountriesNowSpaceApiManager_Tests(SqliteDatabaseFixture fixture)
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
        var newCountry = new cns_Country { Iso2 = "DE", Country = "Germany", Iso3 = "DEU" };
        await _sut.CreateCountryAsync(newCountry);

        var retrieved = await _sut.GetCountryByIso2Async("DE");
        retrieved.Should().NotBeNull();
        retrieved!.Country.Should().Be("Germany");

        await _sut.DeleteCountryByIso2Async("DE");
    }

    // ── DeleteCountryByIso2Async ──────────────────────────────────────────────────

    [Fact]
    public async Task DeleteCountryByIso2Async_UnknownIso2_ReturnsFalse()
    {
        var result = await _sut.DeleteCountryByIso2Async("XX");

        result.Should().BeFalse();
    }

    // ── cns_City ──────────────────────────────────────────────────────────────────

    [Fact]
    public async Task CreateCityAsync_PersistsCity_RetrievableByIso2()
    {
        await _sut.CreateCityAsync(new cns_City { Iso2 = "CA", City = "Toronto" });

        var cities = await _sut.GetCitiesByIso2Async("CA");
        cities.Should().ContainSingle(c => c.City == "Toronto");

        var deleted = await _sut.DeleteCitiesByIso2Async("CA");
        deleted.Should().BeTrue();
    }

    // ── cns_CountryFlag ───────────────────────────────────────────────────────────

    [Fact]
    public async Task CreateCountryFlagAsync_PersistsFlag_RetrievableByIso2()
    {
        await _sut.CreateCountryFlagAsync(new cns_CountryFlag
        {
            Iso2 = "BR",
            Name = "Brazil",
            Flag = "https://example.test/br.svg"
        });

        var flag = await _sut.GetCountryFlagByIso2Async("BR");
        flag.Should().NotBeNull();
        flag!.Flag.Should().Be("https://example.test/br.svg");

        var deleted = await _sut.DeleteCountryFlagByIso2Async("BR");
        deleted.Should().BeTrue();
    }
}
