namespace JBC.ExploreTheWorld.UnitTests.CL;

public class DataResult_Row_Tests
{
    [Fact]
    public void Data_InitializesAsEmptyList()
    {
        var result = new DataResult_Row<string>();

        result.Data.Should().NotBeNull().And.BeEmpty();
    }

    [Fact]
    public void Source_DefaultsToApi()
    {
        var result = new DataResult_Row<string>();

        result.Source.Should().Be(DataSource_Enum.Api);
    }

    [Fact]
    public void Source_CanBeSetToDatabase()
    {
        var result = new DataResult_Row<string> { Source = DataSource_Enum.Database };

        result.Source.Should().Be(DataSource_Enum.Database);
    }

    [Fact]
    public void Data_HoldsAssignedItems()
    {
        var items = new System.Collections.Generic.List<string> { "Australia", "Brazil" };
        var result = new DataResult_Row<string> { Data = items };

        result.Data.Should().BeEquivalentTo(items);
    }
}
