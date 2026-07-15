namespace JBC.ExploreTheWorld.UnitTests.CL;

public class Duration_Helper_Tests
{
    [Theory]
    [InlineData(0, "0.0s")]
    [InlineData(0.05, "0.1s")]
    [InlineData(45.67, "45.7s")]
    [InlineData(120, "2m 0.0s")]
    [InlineData(130, "2m 10.0s")]
    [InlineData(3720, "1h 2m 0.0s")]
    [InlineData(7395.5, "2h 3m 15.5s")]
    public void Format_WritesHoursMinutesSecondsWithMillisecondDecimal(double totalSeconds, string expected)
    {
        Duration_Helper.Format(TimeSpan.FromSeconds(totalSeconds)).Should().Be(expected);
    }
}
