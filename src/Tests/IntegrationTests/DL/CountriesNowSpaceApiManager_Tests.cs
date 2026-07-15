using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;

namespace JBC.ExploreTheWorld.IntegrationTests.DL;

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
        result.Select(c => c.Iso2).Should().Contain(["AU", "BR", "CA"]);
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
        var factory = _fixture.CreateFactory();
        var repo = new CountriesNowSpaceApiManager__Repo(factory);

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

    [Fact]
    public async Task DeleteCitiesByIso2Async_NoCities_ReturnsFalse()
    {
        var result = await _sut.DeleteCitiesByIso2Async("ZZ");

        result.Should().BeFalse();
    }

    // ── cns_CountryCapital ────────────────────────────────────────────────────────

    [Fact]
    public async Task CreateCountryCapitalAsync_PersistsCapital_RetrievableByIso2()
    {
        await _sut.CreateCountryCapitalAsync(new cns_CountryCapital
        {
            Iso2    = "BR",
            Name    = "Brazil",
            Capital = "Brasília"
        });

        var capital = await _sut.GetCountryCapitalByIso2Async("BR");
        capital.Should().NotBeNull();
        capital!.Capital.Should().Be("Brasília");

        var deleted = await _sut.DeleteCountryCapitalByIso2Async("BR");
        deleted.Should().BeTrue();
    }

    [Fact]
    public async Task GetCountryCapitalByIso2Async_UnknownIso2_ReturnsNull()
    {
        var result = await _sut.GetCountryCapitalByIso2Async("ZZ");

        result.Should().BeNull();
    }

    // ── cns_CountryFlag ───────────────────────────────────────────────────────────

    [Fact]
    public async Task CreateCountryFlagAsync_PersistsFlag_RetrievableByIso2()
    {
        await _sut.CreateCountryFlagAsync(new cns_CountryFlag
        {
            Iso2 = "AU",
            Name = "Australia",
            Flag = "https://example.test/au.svg"
        });

        var flag = await _sut.GetCountryFlagByIso2Async("AU");
        flag.Should().NotBeNull();
        flag!.Flag.Should().Be("https://example.test/au.svg");

        var deleted = await _sut.DeleteCountryFlagByIso2Async("AU");
        deleted.Should().BeTrue();
    }

    [Fact]
    public async Task UpdateCountryFlagAsync_ChangesPersistedValues()
    {
        await _sut.CreateCountryFlagAsync(new cns_CountryFlag { Iso2 = "CA", Flag = "old.svg" });

        var flag = await _sut.GetCountryFlagByIso2Async("CA");
        flag!.Flag = "new.svg";
        await _sut.UpdateCountryFlagAsync(flag);

        var updated = await _sut.GetCountryFlagByIso2Async("CA");
        updated!.Flag.Should().Be("new.svg");

        await _sut.DeleteCountryFlagByIso2Async("CA");
    }

    // ── cns_CountryStates ─────────────────────────────────────────────────────────

    [Fact]
    public async Task CreateCountryStatesAsync_PersistsStates_WithChildren()
    {
        var created = await _sut.CreateCountryStatesAsync(new cns_CountryStates
        {
            Name   = "Australia",
            Iso3   = "AUS",
            States = new List<cns_CountryState>
            {
                new() { Name = "New South Wales", StateCode = "NSW" },
                new() { Name = "Victoria",        StateCode = "VIC" }
            }
        });

        var retrieved = await _sut.GetCountryStatesByIdAsync(created.ID);
        retrieved.Should().NotBeNull();
        retrieved!.States.Should().HaveCount(2);

        var deleted = await _sut.DeleteCountryStatesByIdAsync(created.ID);
        deleted.Should().BeTrue();
    }

    [Fact]
    public async Task GetCountryStatesByIdAsync_UnknownId_ReturnsNull()
    {
        var result = await _sut.GetCountryStatesByIdAsync(int.MaxValue);

        result.Should().BeNull();
    }
}
