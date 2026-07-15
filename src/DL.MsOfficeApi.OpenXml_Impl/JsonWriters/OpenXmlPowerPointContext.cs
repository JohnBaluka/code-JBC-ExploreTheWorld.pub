using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Everything needed to resolve effective (inherited) values for shapes on one
    // PresentationML page: the theme + color map of the governing master, the
    // placeholder inheritance chain, the master text styles and the timing tree.
    internal sealed class OpenXmlPowerPointContext
    {
        public OpenXmlPowerPointPageKind Kind { get; private set; }
        public OpenXmlPart Part { get; private set; } = null!;
        public PresentationPart PresentationPart { get; private set; } = null!;
        public SlideLayoutPart? LayoutPart { get; private set; }
        public OpenXmlPart? MasterPart { get; private set; }
        public ThemePart? ThemePart { get; private set; }
        public P.ColorMap? ColorMap { get; private set; }
        public D.OverrideColorMapping? ColorMapOverride { get; private set; }
        public OpenXmlPowerPointTimingInfo Timing { get; private set; } = new OpenXmlPowerPointTimingInfo();

        private OpenXmlPowerPointContext()
        {
        }

        public static OpenXmlPowerPointContext ForSlide(PresentationPart presentationPart, SlidePart slidePart)
        {
            var layoutPart = slidePart.SlideLayoutPart;

            return new OpenXmlPowerPointContext
            {
                Kind = OpenXmlPowerPointPageKind.Slide,
                Part = slidePart,
                PresentationPart = presentationPart,
                LayoutPart = layoutPart,
                MasterPart = layoutPart?.SlideMasterPart,
                ThemePart = layoutPart?.SlideMasterPart?.ThemePart,
                ColorMap = layoutPart?.SlideMasterPart?.SlideMaster?.ColorMap,
                ColorMapOverride = slidePart.Slide?.ColorMapOverride?.GetFirstChild<D.OverrideColorMapping>()
                    ?? layoutPart?.SlideLayout?.ColorMapOverride?.GetFirstChild<D.OverrideColorMapping>(),
                Timing = OpenXmlPowerPointTimingInfo.Parse(slidePart.Slide?.Timing),
            };
        }

        public static OpenXmlPowerPointContext ForLayout(PresentationPart presentationPart, SlideLayoutPart layoutPart)
        {
            return new OpenXmlPowerPointContext
            {
                Kind = OpenXmlPowerPointPageKind.SlideLayout,
                Part = layoutPart,
                PresentationPart = presentationPart,
                LayoutPart = layoutPart,
                MasterPart = layoutPart.SlideMasterPart,
                ThemePart = layoutPart.SlideMasterPart?.ThemePart,
                ColorMap = layoutPart.SlideMasterPart?.SlideMaster?.ColorMap,
                ColorMapOverride = layoutPart.SlideLayout?.ColorMapOverride?.GetFirstChild<D.OverrideColorMapping>(),
                Timing = OpenXmlPowerPointTimingInfo.Parse(layoutPart.SlideLayout?.Timing),
            };
        }

        public static OpenXmlPowerPointContext ForMaster(PresentationPart presentationPart, SlideMasterPart masterPart)
        {
            return new OpenXmlPowerPointContext
            {
                Kind = OpenXmlPowerPointPageKind.SlideMaster,
                Part = masterPart,
                PresentationPart = presentationPart,
                MasterPart = masterPart,
                ThemePart = masterPart.ThemePart,
                ColorMap = masterPart.SlideMaster?.ColorMap,
            };
        }

        public static OpenXmlPowerPointContext ForNotesSlide(PresentationPart presentationPart, NotesSlidePart notesSlidePart)
        {
            var notesMasterPart = presentationPart.NotesMasterPart;

            return new OpenXmlPowerPointContext
            {
                Kind = OpenXmlPowerPointPageKind.NotesSlide,
                Part = notesSlidePart,
                PresentationPart = presentationPart,
                MasterPart = notesMasterPart,
                ThemePart = notesMasterPart?.ThemePart,
                ColorMap = notesMasterPart?.NotesMaster?.ColorMap,
                ColorMapOverride = notesSlidePart.NotesSlide?.ColorMapOverride?.GetFirstChild<D.OverrideColorMapping>(),
            };
        }

        // Context for the notes page COM materializes for slides without a notes part.
        public static OpenXmlPowerPointContext ForSynthesizedNotesSlide(PresentationPart presentationPart)
        {
            var notesMasterPart = presentationPart.NotesMasterPart;

            return new OpenXmlPowerPointContext
            {
                Kind = OpenXmlPowerPointPageKind.NotesSlide,
                Part = (OpenXmlPart?)notesMasterPart ?? presentationPart,
                PresentationPart = presentationPart,
                MasterPart = notesMasterPart,
                ThemePart = notesMasterPart?.ThemePart,
                ColorMap = notesMasterPart?.NotesMaster?.ColorMap,
            };
        }

        public static OpenXmlPowerPointContext ForNotesMaster(PresentationPart presentationPart, NotesMasterPart notesMasterPart)
        {
            return new OpenXmlPowerPointContext
            {
                Kind = OpenXmlPowerPointPageKind.NotesMaster,
                Part = notesMasterPart,
                PresentationPart = presentationPart,
                MasterPart = notesMasterPart,
                ThemePart = notesMasterPart.ThemePart,
                ColorMap = notesMasterPart.NotesMaster?.ColorMap,
            };
        }

        public static OpenXmlPowerPointContext ForHandoutMaster(PresentationPart presentationPart, HandoutMasterPart handoutMasterPart)
        {
            return new OpenXmlPowerPointContext
            {
                Kind = OpenXmlPowerPointPageKind.HandoutMaster,
                Part = handoutMasterPart,
                PresentationPart = presentationPart,
                MasterPart = handoutMasterPart,
                ThemePart = handoutMasterPart.ThemePart,
                ColorMap = handoutMasterPart.HandoutMaster?.ColorMap,
            };
        }

        public bool IsMasterPage
        {
            get
            {
                return Kind == OpenXmlPowerPointPageKind.SlideMaster
                    || Kind == OpenXmlPowerPointPageKind.NotesMaster
                    || Kind == OpenXmlPowerPointPageKind.HandoutMaster;
            }
        }

        // Maps a scheme color name as written in the file (e.g. "bg1", "accent3") to
        // the theme color scheme slot ("lt1", "accent3", ...) through the color map.
        public string MapSchemeColorName(string schemeName)
        {
            if (ColorMapOverride != null && OverrideSlot(ColorMapOverride, schemeName) is string overridden) return overridden;
            if (ColorMap != null && MasterSlot(ColorMap, schemeName) is string mapped) return mapped;

            return schemeName switch
            {
                "bg1" => "lt1",
                "tx1" => "dk1",
                "bg2" => "lt2",
                "tx2" => "dk2",
                _ => schemeName,
            };
        }

        private static string? MasterSlot(P.ColorMap map, string schemeName)
        {
            return schemeName switch
            {
                "bg1" => map.Background1?.InnerText,
                "tx1" => map.Text1?.InnerText,
                "bg2" => map.Background2?.InnerText,
                "tx2" => map.Text2?.InnerText,
                "accent1" => map.Accent1?.InnerText,
                "accent2" => map.Accent2?.InnerText,
                "accent3" => map.Accent3?.InnerText,
                "accent4" => map.Accent4?.InnerText,
                "accent5" => map.Accent5?.InnerText,
                "accent6" => map.Accent6?.InnerText,
                "hlink" => map.Hyperlink?.InnerText,
                "folHlink" => map.FollowedHyperlink?.InnerText,
                _ => null,
            };
        }

        private static string? OverrideSlot(D.OverrideColorMapping map, string schemeName)
        {
            return schemeName switch
            {
                "bg1" => map.Background1?.InnerText,
                "tx1" => map.Text1?.InnerText,
                "bg2" => map.Background2?.InnerText,
                "tx2" => map.Text2?.InnerText,
                "accent1" => map.Accent1?.InnerText,
                "accent2" => map.Accent2?.InnerText,
                "accent3" => map.Accent3?.InnerText,
                "accent4" => map.Accent4?.InnerText,
                "accent5" => map.Accent5?.InnerText,
                "accent6" => map.Accent6?.InnerText,
                "hlink" => map.Hyperlink?.InnerText,
                "folHlink" => map.FollowedHyperlink?.InnerText,
                _ => null,
            };
        }

        // Hex RGB ("RRGGBB") of a theme color scheme slot ("dk1", "accent1", ...).
        public string? GetThemeColorHex(string slotName)
        {
            var scheme = ThemePart?.Theme?.ThemeElements?.ColorScheme;
            if (scheme == null) return null;

            D.Color2Type? color = slotName switch
            {
                "dk1" => scheme.Dark1Color,
                "lt1" => scheme.Light1Color,
                "dk2" => scheme.Dark2Color,
                "lt2" => scheme.Light2Color,
                "accent1" => scheme.Accent1Color,
                "accent2" => scheme.Accent2Color,
                "accent3" => scheme.Accent3Color,
                "accent4" => scheme.Accent4Color,
                "accent5" => scheme.Accent5Color,
                "accent6" => scheme.Accent6Color,
                "hlink" => scheme.Hyperlink,
                "folHlink" => scheme.FollowedHyperlinkColor,
                _ => null,
            };

            return color?.RgbColorModelHex?.Val?.Value ?? color?.SystemColor?.LastColor?.Value;
        }

        public D.FontScheme? FontScheme
        {
            get
            {
                return ThemePart?.Theme?.ThemeElements?.FontScheme;
            }
        }

        public D.FormatScheme? FormatScheme
        {
            get
            {
                return ThemePart?.Theme?.ThemeElements?.FormatScheme;
            }
        }

        // The layout and master shapes a placeholder inherits from, nearest first.
        public List<P.Shape> GetPlaceholderChain(P.PlaceholderShape placeholder)
        {
            var chain = new List<P.Shape>();

            if (Kind == OpenXmlPowerPointPageKind.Slide && LayoutPart != null)
            {
                var layoutShape = FindPlaceholder(LayoutPart.SlideLayout?.CommonSlideData?.ShapeTree, placeholder, matchByIndexFirst: true);
                if (layoutShape != null) chain.Add(layoutShape);

                var layoutPlaceholder = layoutShape == null ? placeholder : PlaceholderOf(layoutShape) ?? placeholder;
                var masterShape = FindMasterPlaceholder(layoutPlaceholder);
                if (masterShape != null) chain.Add(masterShape);
            }
            else if (Kind == OpenXmlPowerPointPageKind.SlideLayout || Kind == OpenXmlPowerPointPageKind.NotesSlide)
            {
                var masterShape = FindMasterPlaceholder(placeholder);
                if (masterShape != null) chain.Add(masterShape);
            }

            return chain;
        }

        private P.Shape? FindMasterPlaceholder(P.PlaceholderShape placeholder)
        {
            var masterTree = MasterPart switch
            {
                SlideMasterPart slideMaster => slideMaster.SlideMaster?.CommonSlideData?.ShapeTree,
                NotesMasterPart notesMaster => notesMaster.NotesMaster?.CommonSlideData?.ShapeTree,
                HandoutMasterPart handoutMaster => handoutMaster.HandoutMaster?.CommonSlideData?.ShapeTree,
                _ => null,
            };

            if (Kind == OpenXmlPowerPointPageKind.NotesSlide)
            {
                return FindPlaceholder(masterTree, placeholder, matchByIndexFirst: false);
            }

            // Slide masters only define title/body/dt/ftr/sldNum placeholders; every
            // other type falls back to the body placeholder.
            string type = placeholder.Type?.InnerText ?? "body";
            string masterType = type switch
            {
                "title" or "ctrTitle" => "title",
                "dt" => "dt",
                "ftr" => "ftr",
                "sldNum" => "sldNum",
                _ => "body",
            };

            return masterTree?.Elements<P.Shape>()
                .FirstOrDefault(s => (PlaceholderOf(s)?.Type?.InnerText ?? "body") == masterType);
        }

        private static P.Shape? FindPlaceholder(P.ShapeTree? tree, P.PlaceholderShape placeholder, bool matchByIndexFirst)
        {
            if (tree == null) return null;

            uint index = placeholder.Index?.Value ?? 0;
            string type = placeholder.Type?.InnerText ?? "body";

            var shapes = tree.Elements<P.Shape>().ToList();

            if (matchByIndexFirst && index != 0)
            {
                var byIndex = shapes.FirstOrDefault(s => (PlaceholderOf(s)?.Index?.Value ?? 0) == index && PlaceholderOf(s) != null);
                if (byIndex != null) return byIndex;
            }

            return shapes.FirstOrDefault(s => PlaceholderOf(s) != null
                    && TypesMatch(PlaceholderOf(s)!.Type?.InnerText ?? "body", type)
                    && (PlaceholderOf(s)!.Index?.Value ?? 0) == index)
                ?? shapes.FirstOrDefault(s => PlaceholderOf(s) != null
                    && TypesMatch(PlaceholderOf(s)!.Type?.InnerText ?? "body", type));
        }

        private static bool TypesMatch(string first, string second)
        {
            static string Normalize(string type)
            {
                return type switch
                {
                    "ctrTitle" => "title",
                    "subTitle" => "body",
                    _ => type,
                };
            }

            return Normalize(first) == Normalize(second);
        }

        public static P.PlaceholderShape? PlaceholderOf(P.Shape shape)
        {
            return shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
        }

        // The master text list style a shape's runs inherit from.
        public OpenXmlCompositeElement? GetMasterTextStyle(P.PlaceholderShape? placeholder)
        {
            if (Kind == OpenXmlPowerPointPageKind.NotesSlide || Kind == OpenXmlPowerPointPageKind.NotesMaster)
            {
                return (MasterPart as NotesMasterPart)?.NotesMaster?.NotesStyle;
            }

            var textStyles = (MasterPart as SlideMasterPart)?.SlideMaster?.TextStyles;
            if (textStyles == null) return null;

            string type = placeholder == null ? "other" : placeholder.Type?.InnerText ?? "body";

            return type switch
            {
                "title" or "ctrTitle" => textStyles.TitleStyle,
                "body" or "subTitle" or "obj" => textStyles.BodyStyle,
                _ => textStyles.OtherStyle,
            };
        }
    }
}
