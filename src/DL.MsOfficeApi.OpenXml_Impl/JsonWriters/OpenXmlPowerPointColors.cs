using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using D = DocumentFormat.OpenXml.Drawing;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters.OpenXmlJsonValues;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Resolves DrawingML color choices (srgbClr/schemeClr/sysClr/prstClr with their
    // tint/shade/lumMod/lumOff transforms) into the PP.ColorFormat values the COM
    // writers report: Type 1 for literal RGB, Type 2 for scheme colors (with the
    // MsoThemeColorIndex and the legacy PpColorSchemeIndex slot).
    internal static class OpenXmlPowerPointColors
    {
        public const long PowerPointCreator = 1347899476; // "PPT " fourcc

        // The default ColorFormat COM reports where no color is defined (white RGB).
        public static PP.ColorFormat White()
        {
            return new PP.ColorFormat
            {
                Brightness = "0",
                Creator = PowerPointCreator,
                ObjectThemeColor = 0,
                RGB = 16777215,
                SchemeColor = 0,
                TintAndShade = "0",
                Type = 1,
            };
        }

        // The "mixed" ColorFormat COM reports for shapes without a fill concept
        // (connectors, graphics).
        public static PP.ColorFormat Mixed()
        {
            return new PP.ColorFormat
            {
                Brightness = "0",
                Creator = PowerPointCreator,
                ObjectThemeColor = -2,
                RGB = 0,
                SchemeColor = -2,
                TintAndShade = "0",
                Type = -2,
            };
        }

        // Resolves the color choice child of `parent` (e.g. an a:solidFill).
        // `placeholderColor` substitutes schemeClr val="phClr" (style matrix references).
        public static PP.ColorFormat? FromParent(OpenXmlElement? parent, OpenXmlPowerPointContext context,
            OpenXmlElement? placeholderColor = null)
        {
            var choice = ColorChoiceOf(parent);
            if (choice == null) return null;

            return FromChoice(choice, context, placeholderColor);
        }

        public static OpenXmlElement? ColorChoiceOf(OpenXmlElement? parent)
        {
            return parent?.ChildElements.FirstOrDefault(c =>
                c is D.RgbColorModelHex
                || c is D.SchemeColor
                || c is D.SystemColor
                || c is D.PresetColor
                || c is D.HslColor
                || c is D.RgbColorModelPercentage);
        }

        public static PP.ColorFormat? FromChoice(OpenXmlElement choice, OpenXmlPowerPointContext context,
            OpenXmlElement? placeholderColor = null)
        {
            string? hex = null;
            int type = 1;
            int? objectThemeColor = 0;
            int? schemeColor = 0;
            OpenXmlElement transformSource = choice;

            switch (choice)
            {
                case D.RgbColorModelHex rgb:
                    hex = rgb.Val?.Value;
                    break;

                case D.SystemColor system:
                    hex = system.LastColor?.Value ?? "000000";
                    break;

                case D.PresetColor preset:
                    hex = HexFromPresetColor(preset.Val?.InnerText);
                    break;

                case D.SchemeColor scheme:
                    string name = scheme.Val?.InnerText ?? string.Empty;
                    if (name == "phClr")
                    {
                        if (placeholderColor == null) return null;

                        var resolved = FromChoice(placeholderColor, context);
                        if (resolved == null) return null;

                        // The phClr transforms (rare) stack on the reference color's own.
                        ApplyTransformList(resolved, scheme, context);
                        return resolved;
                    }

                    type = 2;
                    objectThemeColor = MsoThemeColorIndexFrom(name);
                    schemeColor = LegacySchemeIndexFrom(name);
                    hex = context.GetThemeColorHex(context.MapSchemeColorName(name));
                    break;

                default:
                    return null;
            }

            if (hex == null) return null;

            var format = new PP.ColorFormat
            {
                Brightness = "0",
                Creator = PowerPointCreator,
                ObjectThemeColor = objectThemeColor,
                RGB = OleFromHex(hex),
                SchemeColor = schemeColor,
                TintAndShade = "0",
                Type = type,
            };

            ApplyTransformList(format, transformSource, context);

            return format;
        }

        // Applies tint/shade/lumMod/lumOff children of `source` to the format's RGB
        // and to the reported Brightness/TintAndShade values.
        private static void ApplyTransformList(PP.ColorFormat format, OpenXmlElement source, OpenXmlPowerPointContext context)
        {
            double? lumMod = null;
            double? lumOff = null;

            foreach (var child in source.ChildElements)
            {
                switch (child)
                {
                    case D.Tint tint when tint.Val?.Value is int tintValue:
                        format.RGB = ApplyLinear(format.RGB ?? 0, c => c * (tintValue / 100000.0) + (1.0 - tintValue / 100000.0));
                        format.TintAndShade = SingleString(1.0 - tintValue / 100000.0);
                        break;

                    case D.Shade shade when shade.Val?.Value is int shadeValue:
                        format.RGB = ApplyLinear(format.RGB ?? 0, c => c * (shadeValue / 100000.0));
                        format.TintAndShade = SingleString(-(1.0 - shadeValue / 100000.0));
                        break;

                    case D.LuminanceModulation modulation when modulation.Val?.Value is int modulationValue:
                        lumMod = modulationValue / 100000.0;
                        break;

                    case D.LuminanceOffset offset when offset.Val?.Value is int offsetValue:
                        lumOff = offsetValue / 100000.0;
                        break;
                }
            }

            if (lumMod.HasValue || lumOff.HasValue)
            {
                format.RGB = ApplyLuminance(format.RGB ?? 0, lumMod ?? 1.0, lumOff ?? 0.0);
                format.Brightness = lumOff > 0 ? SingleString(lumOff.Value) : SingleString((lumMod ?? 1.0) - 1.0);
            }
        }

        public static int OleFromHex(string hex)
        {
            long? ole = OleColorFromHex(hex);
            return (int)(ole ?? 0);
        }

        // Transparency (0..1 Single string) from an alpha transform on the color choice.
        public static string TransparencyFrom(OpenXmlElement? parent)
        {
            return AlphaTransparencyFrom(ColorChoiceOf(parent));
        }

        public static string AlphaTransparencyFrom(OpenXmlElement? colorChoice)
        {
            int? alpha = colorChoice?.ChildElements.OfType<D.Alpha>().FirstOrDefault()?.Val?.Value;
            if (alpha == null) return "0";

            return SingleString((100000.0 - alpha.Value) / 100000.0);
        }

        public static int? MsoThemeColorIndexFrom(string schemeName)
        {
            // MsoThemeColorIndex, from the scheme name as written in the file.
            return schemeName switch
            {
                "dk1" => 1,
                "lt1" => 2,
                "dk2" => 3,
                "lt2" => 4,
                "accent1" => 5,
                "accent2" => 6,
                "accent3" => 7,
                "accent4" => 8,
                "accent5" => 9,
                "accent6" => 10,
                "hlink" => 11,
                "folHlink" => 12,
                "tx1" => 13,
                "bg1" => 14,
                "tx2" => 15,
                "bg2" => 16,
                _ => null,
            };
        }

        // The legacy PPT color scheme has eight slots; scheme colors outside them make
        // the COM SchemeColor getter raise (null).
        public static int? LegacySchemeIndexFrom(string schemeName)
        {
            return schemeName switch
            {
                "bg1" => 1, // ppBackground
                "tx1" => 2, // ppForeground
                "bg2" => 3, // ppShadow
                "tx2" => 4, // ppTitle
                "accent1" => 5, // ppFill
                "accent2" => 6, // ppAccent1
                "hlink" => 7, // ppAccent2
                "folHlink" => 8, // ppAccent3
                _ => null,
            };
        }

        private static string? HexFromPresetColor(string? name)
        {
            return name switch
            {
                "black" => "000000",
                "white" => "FFFFFF",
                "red" => "FF0000",
                "green" => "008000",
                "blue" => "0000FF",
                "yellow" => "FFFF00",
                "gray" => "808080",
                _ => null,
            };
        }

        // -- RGB math (both directions use the OLE BGR long) ------------------------

        private static int ApplyLinear(int ole, Func<double, double> transform)
        {
            (double r, double g, double b) = SrgbFromOle(ole);

            r = LinearToSrgb(transform(SrgbToLinear(r)));
            g = LinearToSrgb(transform(SrgbToLinear(g)));
            b = LinearToSrgb(transform(SrgbToLinear(b)));

            return OleFromSrgb(r, g, b);
        }

        private static int ApplyLuminance(int ole, double lumMod, double lumOff)
        {
            (double r, double g, double b) = SrgbFromOle(ole);
            (double h, double s, double l) = HslFromRgb(r, g, b);

            l = Clamp(l * lumMod + lumOff);

            (r, g, b) = RgbFromHsl(h, s, l);
            return OleFromSrgb(r, g, b);
        }

        private static (double R, double G, double B) SrgbFromOle(int ole)
        {
            double r = (ole & 0xFF) / 255.0;
            double g = ((ole >> 8) & 0xFF) / 255.0;
            double b = ((ole >> 16) & 0xFF) / 255.0;
            return (r, g, b);
        }

        private static int OleFromSrgb(double r, double g, double b)
        {
            int red = (int)Math.Round(Clamp(r) * 255.0);
            int green = (int)Math.Round(Clamp(g) * 255.0);
            int blue = (int)Math.Round(Clamp(b) * 255.0);
            return red + (green << 8) + (blue << 16);
        }

        private static double SrgbToLinear(double c)
        {
            return c <= 0.04045 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
        }

        private static double LinearToSrgb(double c)
        {
            return c <= 0.0031308 ? c * 12.92 : 1.055 * Math.Pow(c, 1.0 / 2.4) - 0.055;
        }

        private static double Clamp(double value)
        {
            return value < 0 ? 0 : (value > 1 ? 1 : value);
        }

        private static (double H, double S, double L) HslFromRgb(double r, double g, double b)
        {
            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            double l = (max + min) / 2.0;

            if (max == min) return (0, 0, l);

            double delta = max - min;
            double s = l > 0.5 ? delta / (2.0 - max - min) : delta / (max + min);

            double h;
            if (max == r) h = (g - b) / delta + (g < b ? 6.0 : 0.0);
            else if (max == g) h = (b - r) / delta + 2.0;
            else h = (r - g) / delta + 4.0;
            h /= 6.0;

            return (h, s, l);
        }

        private static (double R, double G, double B) RgbFromHsl(double h, double s, double l)
        {
            if (s == 0) return (l, l, l);

            double q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
            double p = 2.0 * l - q;

            return (HueToRgb(p, q, h + 1.0 / 3.0), HueToRgb(p, q, h), HueToRgb(p, q, h - 1.0 / 3.0));
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1.0;
            if (t > 1) t -= 1.0;
            if (t < 1.0 / 6.0) return p + (q - p) * 6.0 * t;
            if (t < 1.0 / 2.0) return q;
            if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
            return p;
        }
    }
}
