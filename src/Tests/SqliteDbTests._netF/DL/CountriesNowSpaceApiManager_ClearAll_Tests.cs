using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;

namespace JBC.ExploreTheWorld.SqliteDbTests.DL;

// Separate class (and therefore separate fixture database) so clearing all rows
// cannot interfere with the other repository tests.
public class CountriesNowSpaceApiManager_ClearAll_Tests : IClassFixture<SqliteDatabaseFixture>
{
    private readonly CountriesNowSpaceApiManager__Repo _sut;

    public CountriesNowSpaceApiManager_ClearAll_Tests(SqliteDatabaseFixture fixture)
    {
        _sut = new CountriesNowSpaceApiManager__Repo(fixture.CreateFactory());
    }

    [Fact]
    public async Task ClearAllAsync_RemovesAllSeededRows()
    {
        (await _sut.GetCountriesAsync()).Should().NotBeEmpty();

        await _sut.ClearAllAsync();

        (await _sut.GetCountriesAsync()).Should().BeEmpty();
    }
}
