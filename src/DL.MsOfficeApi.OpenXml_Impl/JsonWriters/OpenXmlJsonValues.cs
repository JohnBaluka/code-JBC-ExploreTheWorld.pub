using System;
using System.Globalization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Conversion helpers shared by the OpenXML JSON writers. Values the OpenXML file
    // format cannot provide are written with the MsOfficeUndefined markers.
    internal static class OpenXmlJsonValues
    {
        private const double EmusPerPoint = 12700.0;

        // Doubles pass through a Single and print with 7 significant digits, exactly
        // like the VBA CStr(Single) the gold writers use ("50.61386", "-2.147484E+09").
        // VBA switches to scientific notation when the fixed form needs more than 7
        // digits after the decimal point ("4.244094E-02" instead of "0.04244094").
        public static string SingleString(double value)
        {
            float single = (float)value;
            string text = single.ToString("G7", CultureInfo.InvariantCulture);

            int decimalIndex = text.IndexOf('.');
            if (text.IndexOf('E') < 0 && decimalIndex >= 0 && text.Length - decimalIndex - 1 > 7)
            {
                text = single.ToString("E6", CultureInfo.InvariantCulture);

                // Trim trailing mantissa zeros and shrink the exponent to two digits
                // ("4.244094E-002" -> "4.244094E-02").
                int exponentIndex = text.IndexOf('E');
                string mantissa = text.Substring(0, exponentIndex).TrimEnd('0').TrimEnd('.');
                string sign = text.Substring(exponentIndex + 1, 1);
                string digits = text.Substring(exponentIndex + 2).TrimStart('0');
                if (digits.Length < 2) digits = digits.PadLeft(2, '0');

                text = mantissa + "E" + sign + digits;
            }

            return text;
        }

        // EMU to points, formatted like the VBA WritePropValueSingleString helper.
        public static string PointsFromEmu(long? emu)
        {
            if (!emu.HasValue) return MsOffice.MsOfficeUndefined.String;

            return SingleString(emu.Value / EmusPerPoint);
        }

        // Twips (twentieths of a point) to points, formatted as a string.
        public static string PointsFromTwips(double? twips)
        {
            if (!twips.HasValue) return MsOffice.MsOfficeUndefined.String;

            return SingleString(twips.Value / 20.0);
        }

        // 60000ths of a degree to degrees, formatted as a string.
        public static string DegreesFrom60000ths(int? value)
        {
            if (!value.HasValue) return "0";

            return SingleString(value.Value / 60000.0);
        }

        // "RRGGBB" hex to the COM OLE color long (BGR byte order) the VBA writers emit.
        public static long? OleColorFromHex(string? rrggbb)
        {
            if (string.IsNullOrEmpty(rrggbb) || rrggbb!.Length < 6) return null;

            try
            {
                int r = Convert.ToInt32(rrggbb.Substring(0, 2), 16);
                int g = Convert.ToInt32(rrggbb.Substring(2, 2), 16);
                int b = Convert.ToInt32(rrggbb.Substring(4, 2), 16);

                return r + (g << 8) + (b << 16);
            }
            catch
            {
                return null;
            }
        }

        // Half-points (Office font sizes) to points, formatted as a string.
        public static string PointsFromHalfPoints(double? halfPoints)
        {
            if (!halfPoints.HasValue) return MsOffice.MsOfficeUndefined.String;

            return SingleString(halfPoints.Value / 2.0);
        }

        // Hundredths of a point (DrawingML font sizes) to points, formatted as a string.
        public static string PointsFromHundredths(double? hundredths)
        {
            if (!hundredths.HasValue) return MsOffice.MsOfficeUndefined.String;

            return SingleString(hundredths.Value / 100.0);
        }

        // MsoTriState from a boolean (msoTrue = -1, msoFalse = 0).
        public static int TriState(bool value)
        {
            return value ? -1 : 0;
        }

        public static string ExtensionFromContentType(string? contentType)
        {
            return contentType switch
            {
                "image/png" => "png",
                "image/jpeg" => "jpg",
                "image/gif" => "gif",
                "image/bmp" => "bmp",
                "image/tiff" => "tif",
                "image/svg+xml" => "svg",
                "image/x-emf" => "emf",
                "image/x-wmf" => "wmf",
                _ => "bin",
            };
        }
    }
}
