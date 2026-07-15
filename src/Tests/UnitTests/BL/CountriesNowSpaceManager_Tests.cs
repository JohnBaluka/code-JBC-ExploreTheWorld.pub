using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.UnitTests.BL;

public class CountriesNowSpaceManager_Tests
{
    private readonly Mock<CountriesNowSpaceApi_Interface> _mockApi;
    private readonly Mock<CountriesNowSpaceApiManager__Repo__Interface> _mockDbManager;
    private readonly CountriesNowSpaceManager__Service _sut;

    public CountriesNowSpaceManager_Tests()
    {
        _mockApi = new Mock<CountriesNowSpaceApi_Interface>();
        _mockDbManager = new Mock<CountriesNowSpaceApiManager__Repo__Interface>();
        _sut = new CountriesNowSpaceManager__Service(_mockApi.Object, _mockDbManager.Object);
    }

    // ── GetAllCountriesAsync ──────────────────────────────────────────────────────

    [Fact]
    public async Task GetAllCountriesAsync_WhenDbHasData_ReturnsDatabaseSource()
    {
        var dbCountries = new System.Collections.Generic.List<cns_Country>
        {
            new() { Country = "Australia", Iso2 = "AU", Iso3 = "AUS" },
            new() { Country = "Brazil",    Iso2 = "BR", Iso3 = "BRA" }
        };

        _mockDbManager
            .Setup(db => db.GetCountriesAsync())
            .ReturnsAsync(dbCountries);

        var result = await _sut.GetAllCountriesAsync();

        result.Source.Should().Be(DataSource_Enum.Database);
        result.Data.Should().HaveCount(2);
        _mockApi.Verify(api => api.GetAllCountriesAsync(), Times.Never);
    }

    [Fact]
    public async Task GetAllCountriesAsync_WhenDbIsEmpty_CallsApiAndReturnsApiSource()
    {
        _mockDbManager
            .Setup(db => db.GetCountriesAsync())
            .ReturnsAsync(new System.Collections.Generic.List<cns_Country>());

        var apiCountries = new System.Collections.Generic.List<CountryBasic_Row>
        {
            new() { Country = "Australia", Iso2 = "AU", Iso3 = "AUS" }
        };

        _mockApi
            .Setup(api => api.GetAllCountriesAsync())
            .ReturnsAsync(apiCountries);

        _mockDbManager
            .Setup(db => db.CreateCountryAsync(It.IsAny<cns_Country>()))
            .ReturnsAsync(new cns_Country());

        var result = await _sut.GetAllCountriesAsync();

        result.Source.Should().Be(DataSource_Enum.Api);
        result.Data.Should().HaveCount(1);
        _mockDbManager.Verify(db => db.CreateCountryAsync(It.IsAny<cns_Country>()), Times.Once);
    }

    [Fact]
    public async Task GetAllCountriesAsync_WhenApiReturnsRowWithoutIso2_DoesNotPersistToDb()
    {
        _mockDbManager
            .Setup(db => db.GetCountriesAsync())
            .ReturnsAsync(new System.Collections.Generic.List<cns_Country>());

        var apiCountries = new System.Collections.Generic.List<CountryBasic_Row>
        {
            new() { Country = "Unknown", Iso2 = null, Iso3 = "UNK" }
        };

        _mockApi
            .Setup(api => api.GetAllCountriesAsync())
            .ReturnsAsync(apiCountries);

        await _sut.GetAllCountriesAsync();

        _mockDbManager.Verify(db => db.CreateCountryAsync(It.IsAny<cns_Country>()), Times.Never);
    }

    // ── GetCountryStatesAsync ─────────────────────────────────────────────────────

    [Fact]
    public async Task GetCountryStatesAsync_WhenDbHasCountry_ReturnsDatabaseSource()
    {
        var dbStates = new System.Collections.Generic.List<cns_CountryStates>
        {
            new()
            {
                Name   = "Australia",
                Iso3   = "AUS",
                States = new System.Collections.Generic.List<cns_CountryState>
                {
                    new() { Name = "New South Wales", StateCode = "NSW" },
                    new() { Name = "Victoria",        StateCode = "VIC" }
                }
            }
        };

        _mockDbManager
            .Setup(db => db.GetCountryStatesAsync())
            .ReturnsAsync(dbStates);

        var result = await _sut.GetCountryStatesAsync("Australia");

        result.Source.Should().Be(DataSource_Enum.Database);
        result.Data.Should().HaveCount(2);
        _mockApi.Verify(api => api.GetCountryStatesAsync(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public async Task GetCountryStatesAsync_MatchesDbCountry_CaseInsensitively()
    {
        var dbStates = new System.Collections.Generic.List<cns_CountryStates>
        {
            new()
            {
                Name   = "Australia",
                States = new System.Collections.Generic.List<cns_CountryState>
                {
                    new() { Name = "Victoria", StateCode = "VIC" }
                }
            }
        };

        _mockDbManager
            .Setup(db => db.GetCountryStatesAsync())
            .ReturnsAsync(dbStates);

        var result = await _sut.GetCountryStatesAsync("AUSTRALIA");

        result.Source.Should().Be(DataSource_Enum.Database);
        result.Data.Should().ContainSingle(s => s.StateCode == "VIC");
    }

    [Fact]
    public async Task GetCountryStatesAsync_WhenDbIsEmpty_CallsApiAndPersists()
    {
        _mockDbManager
            .Setup(db => db.GetCountryStatesAsync())
            .ReturnsAsync(new System.Collections.Generic.List<cns_CountryStates>());

        _mockApi
            .Setup(api => api.GetCountryStatesAsync("Brazil"))
            .ReturnsAsync(new CountryStates_Row
            {
                Name   = "Brazil",
                Iso3   = "BRA",
                States = new System.Collections.Generic.List<CountryState_Row>
                {
                    new() { Name = "São Paulo", StateCode = "SP" }
                }
            });

        _mockDbManager
            .Setup(db => db.CreateCountryStatesAsync(It.IsAny<cns_CountryStates>()))
            .ReturnsAsync(new cns_CountryStates());

        var result = await _sut.GetCountryStatesAsync("Brazil");

        result.Source.Should().Be(DataSource_Enum.Api);
        result.Data.Should().ContainSingle(s => s.StateCode == "SP");
        _mockDbManager.Verify(db => db.CreateCountryStatesAsync(It.IsAny<cns_CountryStates>()), Times.Once);
    }

    [Fact]
    public async Task GetCountryStatesAsync_WhenApiReturnsNull_ReturnsEmptyAndDoesNotPersist()
    {
        _mockDbManager
            .Setup(db => db.GetCountryStatesAsync())
            .ReturnsAsync(new System.Collections.Generic.List<cns_CountryStates>());

        _mockApi
            .Setup(api => api.GetCountryStatesAsync(It.IsAny<string>()))
            .ReturnsAsync((CountryStates_Row?)null);

        var result = await _sut.GetCountryStatesAsync("Atlantis");

        result.Source.Should().Be(DataSource_Enum.Api);
        result.Data.Should().BeEmpty();
        _mockDbManager.Verify(db => db.CreateCountryStatesAsync(It.IsAny<cns_CountryStates>()), Times.Never);
    }

    // ── GetCountryFlagsAsync ──────────────────────────────────────────────────────

    [Fact]
    public async Task GetCountryFlagsAsync_WhenDbHasData_ReturnsDatabaseSource()
    {
        var dbFlags = new System.Collections.Generic.List<cns_CountryFlag>
        {
            new() { Iso2 = "AU", Name = "Australia", Flag = "https://example.org/au.svg" },
            new() { Iso2 = "BR", Name = "Brazil",    Flag = "https://example.org/br.svg" }
        };

        _mockDbManager
            .Setup(db => db.GetCountryFlagsAsync())
            .ReturnsAsync(dbFlags);

        var result = await _sut.GetCountryFlagsAsync();

        result.Source.Should().Be(DataSource_Enum.Database);
        result.Data.Should().HaveCount(2);
        _mockApi.Verify(api => api.GetCountryFlagsAsync(), Times.Never);
    }

    [Fact]
    public async Task GetCountryFlagsAsync_WhenDbIsEmpty_CallsApiAndPersists()
    {
        _mockDbManager
            .Setup(db => db.GetCountryFlagsAsync())
            .ReturnsAsync(new System.Collections.Generic.List<cns_CountryFlag>());

        var apiFlags = new System.Collections.Generic.List<CountryFlag_Row>
        {
            new() { Iso2 = "AU", Name = "Australia", Flag = "https://example.org/au.svg" }
        };

        _mockApi
            .Setup(api => api.GetCountryFlagsAsync())
            .ReturnsAsync(apiFlags);

        _mockDbManager
            .Setup(db => db.CreateCountryFlagAsync(It.IsAny<cns_CountryFlag>()))
            .ReturnsAsync(new cns_CountryFlag());

        var result = await _sut.GetCountryFlagsAsync();

        result.Source.Should().Be(DataSource_Enum.Api);
        result.Data.Should().HaveCount(1);
        _mockDbManager.Verify(db => db.CreateCountryFlagAsync(It.IsAny<cns_CountryFlag>()), Times.Once);
    }

    [Fact]
    public async Task GetCountryFlagsAsync_SkipsRowsWithoutIso2_AndDuplicateIso2()
    {
        _mockDbManager
            .Setup(db => db.GetCountryFlagsAsync())
            .ReturnsAsync(new System.Collections.Generic.List<cns_CountryFlag>());

        var apiFlags = new System.Collections.Generic.List<CountryFlag_Row>
        {
            new() { Iso2 = null, Name = "No code" },
            new() { Iso2 = "AU", Name = "Australia" },
            new() { Iso2 = "au", Name = "Australia duplicate" }
        };

        _mockApi
            .Setup(api => api.GetCountryFlagsAsync())
            .ReturnsAsync(apiFlags);

        _mockDbManager
            .Setup(db => db.CreateCountryFlagAsync(It.IsAny<cns_CountryFlag>()))
            .ReturnsAsync(new cns_CountryFlag());

        await _sut.GetCountryFlagsAsync();

        _mockDbManager.Verify(db => db.CreateCountryFlagAsync(It.IsAny<cns_CountryFlag>()), Times.Once);
    }

    // ── ClearAllDataAsync ─────────────────────────────────────────────────────────

    [Fact]
    public async Task ClearAllDataAsync_DelegatesTo_DbManagerClearAllAsync()
    {
        _mockDbManager.Setup(db => db.ClearAllAsync()).Returns(Task.CompletedTask);

        await _sut.ClearAllDataAsync();

        _mockDbManager.Verify(db => db.ClearAllAsync(), Times.Once);
    }
}
