namespace JBC.ExploreTheWorld.UnitTests.CL;

public class _Row_Tests
{
    [Fact]
    public void GUID_DefaultsToEmptyGuid()
    {
        var row = new _Row();

        row.GUID.Should().Be(Guid.Empty);
    }

    [Fact]
    public void GetPrimaryKeyValue_ReturnsGuidAsString()
    {
        var id = Guid.NewGuid();
        var row = new _Row { GUID = id };

        row.GetPrimaryKeyValue().Should().Be(id.ToString());
    }

    [Fact]
    public void GetBusinessKeyValue_ThrowsNotImplementedException()
    {
        var row = new _Row();

        var act = () => row.GetBusinessKeyValue();

        act.Should().Throw<NotImplementedException>();
    }
}
