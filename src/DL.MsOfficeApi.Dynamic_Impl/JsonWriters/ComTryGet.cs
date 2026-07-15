using System;
using System.Globalization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters
{
    // Reads values from late-bound COM objects with the same semantics as the VBA
    // JSON writers: a property that throws is written as null. Values are formatted
    // with the invariant culture so the output matches the canonical serializer.
    internal static class ComTryGet
    {
        public static string? Str(Func<object?> getter)
        {
            try
            {
                object? value = getter();
                if (value == null) return string.Empty;
                // Matches VBA CStr semantics: converting an array (e.g. Shape.Vertices)
                // raises a type-mismatch error, which the writers serialize as null, and
                // dates format as "7/4/2026 7:16:25 AM" (US locale, no leading zeros).
                if (value is Array) return null;
                if (value is DateTime dateTime) return dateTime.ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture);
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return null;
            }
        }

        public static int? Int(Func<object?> getter)
        {
            try
            {
                object? value = getter();
                return value == null ? null : Convert.ToInt32(value);
            }
            catch
            {
                return null;
            }
        }

        public static long? Long(Func<object?> getter)
        {
            try
            {
                object? value = getter();
                return value == null ? null : Convert.ToInt64(value);
            }
            catch
            {
                return null;
            }
        }

        public static bool? Bool(Func<object?> getter)
        {
            try
            {
                object? value = getter();
                return value == null ? null : Convert.ToBoolean(value);
            }
            catch
            {
                return null;
            }
        }

        // Matches the VBA WritePropValueSingleString helper (a Single written as a string).
        // VBA CStr(Single) formats with at most 7 significant digits and switches to
        // scientific notation when the fixed form needs more than 7 displayed digits
        // (e.g. 0.04244094 -> "4.244094E-02"), so neither the shortest round-trippable
        // default nor plain "G7" matches the canonical output.
        public static string? SingleStr(Func<object?> getter)
        {
            try
            {
                object? value = getter();
                return value == null ? null : FormatSingleLikeCStr(Convert.ToSingle(value));
            }
            catch
            {
                return null;
            }
        }

        private static string FormatSingleLikeCStr(float value)
        {
            string text = value.ToString("G7", CultureInfo.InvariantCulture);
            if (text.IndexOf('E') >= 0) return text;

            int digits = 0;
            foreach (char c in text)
            {
                if (char.IsDigit(c)) digits++;
            }
            if (text.IndexOf('.') >= 0 && Math.Abs(value) < 1) digits--; // the leading "0." zero

            return digits > 7 ? value.ToString("0.######E+00", CultureInfo.InvariantCulture) : text;
        }

        // Matches the VBA WritePropValueDate format: "mm/dd/yyyy hh:nn:ss AM/PM".
        public static string? DateStr(Func<object?> getter)
        {
            try
            {
                object? value = getter();
                return value == null ? null : Convert.ToDateTime(value).ToString("MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
            }
        }

        public static T? Obj<T>(Func<T?> builder) where T : class
        {
            try
            {
                return builder();
            }
            catch
            {
                return null;
            }
        }
    }
}
