using System;
using System.Globalization;

namespace JBC.ExploreTheWorld.CL
{
    /// <summary>
    /// Formats elapsed durations for log output as hours, minutes, and seconds with the
    /// milliseconds as one decimal on the seconds. Leading units are omitted when zero:
    /// 45.67s → "45.7s", 120s → "2m 0.0s", 130s → "2m 10.0s", 62 minutes → "1h 2m 0.0s".
    /// </summary>
    public static class Duration_Helper
    {
        public static string Format(TimeSpan duration)
        {
            var seconds = duration.Seconds + (duration.Milliseconds / 1000d);
            var secondsText = seconds.ToString("0.0", CultureInfo.InvariantCulture);

            if (duration.TotalHours >= 1)
                return $"{(int)duration.TotalHours}h {duration.Minutes}m {secondsText}s";

            if (duration.TotalMinutes >= 1)
                return $"{duration.Minutes}m {secondsText}s";

            return secondsText + "s";
        }
    }
}
