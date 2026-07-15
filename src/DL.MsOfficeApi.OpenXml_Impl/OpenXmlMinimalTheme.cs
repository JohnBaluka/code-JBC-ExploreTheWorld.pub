using D = DocumentFormat.OpenXml.Drawing;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl
{
    // Minimal ECMA-376 theme for generated presentations. PowerPoint refuses to open a
    // package whose slide master has no theme part, so every exporter attaches this one.
    internal static class OpenXmlMinimalTheme
    {
        public static D.Theme Create()
        {
            return new D.Theme(
                new D.ThemeElements(
                    new D.ColorScheme(
                        new D.Dark1Color(new D.SystemColor { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
                        new D.Light1Color(new D.SystemColor { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
                        new D.Dark2Color(new D.RgbColorModelHex { Val = "44546A" }),
                        new D.Light2Color(new D.RgbColorModelHex { Val = "E7E6E6" }),
                        new D.Accent1Color(new D.RgbColorModelHex { Val = "4472C4" }),
                        new D.Accent2Color(new D.RgbColorModelHex { Val = "ED7D31" }),
                        new D.Accent3Color(new D.RgbColorModelHex { Val = "A5A5A5" }),
                        new D.Accent4Color(new D.RgbColorModelHex { Val = "FFC000" }),
                        new D.Accent5Color(new D.RgbColorModelHex { Val = "5B9BD5" }),
                        new D.Accent6Color(new D.RgbColorModelHex { Val = "70AD47" }),
                        new D.Hyperlink(new D.RgbColorModelHex { Val = "0563C1" }),
                        new D.FollowedHyperlinkColor(new D.RgbColorModelHex { Val = "954F72" }))
                    { Name = "Office" },
                    new D.FontScheme(
                        new D.MajorFont(
                            new D.LatinFont { Typeface = "Calibri Light" },
                            new D.EastAsianFont { Typeface = "" },
                            new D.ComplexScriptFont { Typeface = "" }),
                        new D.MinorFont(
                            new D.LatinFont { Typeface = "Calibri" },
                            new D.EastAsianFont { Typeface = "" },
                            new D.ComplexScriptFont { Typeface = "" }))
                    { Name = "Office" },
                    new D.FormatScheme(
                        new D.FillStyleList(
                            new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }),
                            new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }),
                            new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })),
                        new D.LineStyleList(
                            new D.Outline(new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })) { Width = 6350 },
                            new D.Outline(new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })) { Width = 12700 },
                            new D.Outline(new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })) { Width = 19050 }),
                        new D.EffectStyleList(
                            new D.EffectStyle(new D.EffectList()),
                            new D.EffectStyle(new D.EffectList()),
                            new D.EffectStyle(new D.EffectList())),
                        new D.BackgroundFillStyleList(
                            new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }),
                            new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }),
                            new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })))
                    { Name = "Office" }))
            { Name = "Office Theme" };
        }
    }
}
