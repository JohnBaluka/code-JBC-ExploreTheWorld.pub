using System.ComponentModel;

namespace JBC.ExploreTheWorld.UnitTests.CL;

public class Enum_Extensions_Tests
{
    private enum Annotated_Enum
    {
        [Description("Microsoft SQL Server database")]
        SqlServerDb,

        PlainValue
    }

    [Flags]
    private enum Flag_Enum : uint
    {
        None   = 0,
        First  = 1,
        Second = 2
    }

    // ── GetEnumDisplayName ────────────────────────────────────────────────────────

    [Fact]
    public void GetEnumDisplayName_WithoutAttribute_FallsBackToValueName()
    {
        // DisplayNameAttribute cannot be applied to enum fields, so the fallback
        // (the value name itself) is the only reachable path for plain enums.
        Enum_Extensions.GetEnumDisplayName(Annotated_Enum.PlainValue).Should().Be("PlainValue");
    }

    // ── GetEnumDescription ────────────────────────────────────────────────────────

    [Fact]
    public void GetEnumDescription_WithDescriptionAttribute_ReturnsAttributeValue()
    {
        Enum_Extensions.GetEnumDescription(Annotated_Enum.SqlServerDb)
            .Should().Be("Microsoft SQL Server database");
    }

    [Fact]
    public void GetEnumDescription_WithoutAttribute_FallsBackToValueName()
    {
        Enum_Extensions.GetEnumDescription(Annotated_Enum.PlainValue).Should().Be("PlainValue");
    }

    // ── AsUpperCamelCaseName ──────────────────────────────────────────────────────

    [Fact]
    public void AsUpperCamelCaseName_SplitsUpperCamelCaseIntoWords()
    {
        Annotated_Enum.PlainValue.AsUpperCamelCaseName().Should().Be("Plain Value");
    }

    [Fact]
    public void AsUpperCamelCaseName_SingleWord_IsUnchanged()
    {
        DataSource_Enum.Database.AsUpperCamelCaseName().Should().Be("Database");
    }

    // ── EnumToList ────────────────────────────────────────────────────────────────

    [Fact]
    public void EnumToList_ReturnsAllValues()
    {
        var values = Enum_Extensions.EnumToList<DataSource_Enum>();

        values.Should().Contain(DataSource_Enum.Api).And.Contain(DataSource_Enum.Database);
    }

    [Fact]
    public void EnumToList_NonEnumType_Throws()
    {
        var act = () => Enum_Extensions.EnumToList<int>();

        act.Should().Throw<ArgumentException>();
    }

    // ── IsSet ─────────────────────────────────────────────────────────────────────

    [Fact]
    public void IsSet_FlagPresent_ReturnsTrue()
    {
        var value = Flag_Enum.First | Flag_Enum.Second;

        value.IsSet(Flag_Enum.First).Should().BeTrue();
    }

    [Fact]
    public void IsSet_FlagAbsent_ReturnsFalse()
    {
        Flag_Enum.First.IsSet(Flag_Enum.Second).Should().BeFalse();
    }
}
