using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters.OpenXmlJsonValues;
using Undefined = JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice.MsOfficeUndefined;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint entity graph from a .pptx
    // file and writes it with MsOfficeJsonSerializer. The schema matches
    // MSO_MsPowerPoint_JsonWriter.bas; VBA properties the OpenXML file format cannot
    // provide are written as "**Undefined" (strings), -99 (enums/numerics) or null
    // (values whose COM getter raises for that object kind).
    public static class MsPowerPointJsonWriter
    {
        private const int msoAutoShape = 1;
        private const int msoChart = 3;
        private const int msoFreeform = 5;
        private const int msoGroup = 6;
        private const int msoPicture = 13;
        private const int msoPlaceholder = 14;
        private const int msoMedia = 16;
        private const int msoTextBox = 17;
        private const int msoTable = 19;
        private const int msoSmartArt = 24;
        private const int msoGraphic = 28;

        private const long PptCreator = OpenXmlPowerPointColors.PowerPointCreator;

        // MsoAutoShapeType for shapes whose geometry is not a preset autoshape.
        private const int msoShapeNotPrimitive = 138;
        private const int msoShapeMixed = -2;

        // CStr(CSng(...)) of the not-set sentinel PowerPoint stores for AdvanceTime.
        private const string AdvanceTimeNotSet = "9.223372E+15";
        private const string SingleMixed = "-2.147484E+09";

        // Preset geometry names mapped to the MsoAutoShapeType values the COM writers report.
        private static readonly Dictionary<string, int> MsoAutoShapeTypes = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            ["rect"] = 1, ["parallelogram"] = 2, ["trapezoid"] = 3, ["diamond"] = 4, ["roundRect"] = 5,
            ["octagon"] = 6, ["triangle"] = 7, ["rtTriangle"] = 8, ["ellipse"] = 9, ["hexagon"] = 10,
            ["plus"] = 11, ["pentagon"] = 12, ["can"] = 13, ["cube"] = 14, ["bevel"] = 15,
            ["foldedCorner"] = 16, ["smileyFace"] = 17, ["donut"] = 18, ["noSmoking"] = 19,
            ["blockArc"] = 20, ["heart"] = 21, ["lightningBolt"] = 22, ["sun"] = 23, ["moon"] = 24,
            ["arc"] = 25, ["bracketPair"] = 26, ["bracePair"] = 27, ["plaque"] = 28,
            ["leftBracket"] = 29, ["rightBracket"] = 30, ["leftBrace"] = 31, ["rightBrace"] = 32,
            ["rightArrow"] = 33, ["leftArrow"] = 34, ["upArrow"] = 35, ["downArrow"] = 36,
            ["leftRightArrow"] = 37, ["upDownArrow"] = 38, ["quadArrow"] = 39, ["leftRightUpArrow"] = 40,
            ["bentArrow"] = 41, ["uturnArrow"] = 42, ["leftUpArrow"] = 43, ["bentUpArrow"] = 44,
            ["curvedRightArrow"] = 45, ["curvedLeftArrow"] = 46, ["curvedUpArrow"] = 47,
            ["curvedDownArrow"] = 48, ["stripedRightArrow"] = 49, ["notchedRightArrow"] = 50,
            ["homePlate"] = 51, ["chevron"] = 52,
            ["wedgeRectCallout"] = 105, ["wedgeRoundRectCallout"] = 106, ["wedgeEllipseCallout"] = 107,
            ["cloudCallout"] = 108,
        };

        // Connection site counts from the ECMA-376 preset geometry definitions (cxnLst).
        private static readonly Dictionary<string, long> PresetConnectionSites = new Dictionary<string, long>(StringComparer.Ordinal)
        {
            ["rect"] = 4, ["roundRect"] = 4, ["leftBrace"] = 3, ["rightBrace"] = 3,
            ["wedgeRoundRectCallout"] = 5, ["wedgeRectCallout"] = 5,
            ["straightConnector1"] = 0, ["bentConnector2"] = 0, ["bentConnector3"] = 0,
            ["bentConnector4"] = 0, ["bentConnector5"] = 0, ["curvedConnector2"] = 0,
            ["curvedConnector3"] = 0, ["curvedConnector4"] = 0, ["curvedConnector5"] = 0,
            ["line"] = 2,
        };

        // PowerPoint's UI renamed default shape names; the COM Shape.Name reports the
        // legacy name for the built-in "<name> <n>" patterns stored by newer versions.
        private static readonly Dictionary<string, string> LegacyDefaultShapeNames = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Rectangle: Rounded Corners"] = "Rounded Rectangle",
            ["Speech Bubble: Rectangle with Corners Rounded"] = "Rounded Rectangular Callout",
            ["Brace: Left"] = "Left Brace",
            ["Brace: Right"] = "Right Brace",
            ["Freeform: Shape"] = "Freeform",
            ["Freeform: Scribble"] = "Scribble",
            ["Connector: Straight"] = "Straight Connector",
            ["Connector: Elbow"] = "Elbow Connector",
        };

        public static void WritePresentationFileToJson(string sourcePath, string outputJsonPath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            options ??= new MsOfficeJsonWriterOptions();

            log?.Invoke($"Writing (OpenXml): {outputJsonPath}");

            using var document = PresentationDocument.Open(sourcePath, false);

            PP.Presentation entity = BuildPresentation(document, sourcePath, outputJsonPath, options, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputJsonPath);

            if (options.SaveOriginalMedia)
            {
                SaveOriginalMedia(document, outputJsonPath, options, log);
            }

            log?.Invoke($"Done: {outputJsonPath}");
        }

        // Saves every unique image part's original bytes to "{blobFolder}/images/",
        // keeping the package file names (image1.jpeg, ...). Skipped silently when no
        // writable file system exists (in-browser scenarios).
        private static void SaveOriginalMedia(PresentationDocument document, string outputJsonPath,
            MsOfficeJsonWriterOptions options, Action<string>? log)
        {
            try
            {
                var visited = new HashSet<OpenXmlPart>();
                var imagesByUri = new Dictionary<string, ImagePart>(StringComparer.OrdinalIgnoreCase);

                void Walk(OpenXmlPart part)
                {
                    if (!visited.Add(part)) return;

                    if (part is ImagePart image)
                    {
                        imagesByUri[image.Uri.OriginalString] = image;
                    }

                    foreach (var child in part.Parts)
                    {
                        Walk(child.OpenXmlPart);
                    }
                }

                foreach (var pair in document.Parts)
                {
                    Walk(pair.OpenXmlPart);
                }

                if (imagesByUri.Count == 0) return;

                string folder = Path.Combine(options.GetBlobFolderPath(outputJsonPath), "images");
                Directory.CreateDirectory(folder);

                foreach (var image in imagesByUri.Values)
                {
                    string fileName = Path.GetFileName(image.Uri.OriginalString);

                    using var stream = image.GetStream();
                    using var file = File.Create(Path.Combine(folder, fileName));
                    stream.CopyTo(file);
                }

                log?.Invoke($"  Saved {imagesByUri.Count} original media files.");
            }
            catch
            {
                // Original media files are best-effort extras.
            }
        }

        private static PP.Presentation BuildPresentation(PresentationDocument document, string sourcePath,
            string outputJsonPath, MsOfficeJsonWriterOptions options, Action<string>? log)
        {
            var prsPart = document.PresentationPart
                ?? throw new InvalidOperationException("Cannot read presentation part.");

            var prs = new PP.Presentation();

            // Entities
            var masterPart = FirstMasterPart(prsPart);
            prs.SlideMaster = masterPart == null ? null : BuildMaster(prsPart, masterPart, outputJsonPath, options);
            prs.TitleMaster = null;
            prs.HandoutMaster = BuildHandoutMaster(prsPart, outputJsonPath, options);
            prs.NotesMaster = BuildNotesMaster(prsPart, outputJsonPath, options);
            prs.Coauthoring = new PP.Coauthoring
            {
                // A file opened from disk is not in a co-authoring merge session.
                FavorServerEditsDuringMerge = true,
                MergeMode = false,
                PendingUpdates = false,
            };
            prs.PageSetup = BuildPageSetup(prsPart);
            prs.PrintOptions = BuildPrintOptions(prsPart);
            prs.SlideShowSettings = BuildSlideShowSettings(prsPart);

            // Lists
            prs.BuiltInDocumentProperties = BuildBuiltInDocumentProperties(document);
            prs.CustomDocumentProperties = BuildCustomDocumentProperties(document);
            prs.ContentTypeProperties = null;
            prs.ColorSchemes = new List<PP.ColorScheme>();
            prs.CustomXmlParts = BuildCustomXmlParts(document, prsPart);
            prs.Designs = BuildDesigns(prsPart);
            prs.ExtraColors = BuildExtraColors(prsPart);
            prs.Fonts = BuildFonts(prsPart);
            prs.Guides = BuildGuides(prsPart);
            prs.SectionProperties = BuildSectionProperties(prsPart);
            prs.Slides = BuildSlides(prsPart, outputJsonPath, options, log);
            prs.Tags = BuildTags(prsPart.Presentation?.GetFirstChild<P.CustomerDataList>(), prsPart) ?? new List<PP.Tag>();

            // Fields
            prs.Name = Path.GetFileName(sourcePath);
            prs.Path = Path.GetDirectoryName(sourcePath) ?? string.Empty;
            prs.FullName = sourcePath;
            prs.AutoSaveOn = null;
            prs.ChartDataPointTrack = true; // PowerPoint default; no override attribute in this schema
            prs.CreateVideoStatus = 0; // ppMediaTaskStatusNone: no video export task exists for a file on disk
            prs.DefaultLanguageID = DefaultLanguageIdFrom(prsPart);
            prs.DisplayComments = -1;
            prs.EncryptionProvider = string.Empty;
            prs.EnvelopeVisible = 0;
            prs.FarEastLineBreakLanguage = null;
            prs.FarEastLineBreakLevel = 1; // ppFarEastLineBreakLevelNormal (schema default)
            prs.Final = OpenXmlDocumentProperties.GetCustomPropertyValue(document.CustomFilePropertiesPart, "_MarkAsFinal")
                is string markAsFinal
                ? string.Equals(markAsFinal, "true", StringComparison.OrdinalIgnoreCase) || markAsFinal == "1"
                : false;
            prs.GridDistance = GridDistanceFrom(prsPart);
            prs.HasHandoutMaster = TriState(prsPart.HandoutMasterPart != null);
            prs.HasNotesMaster = TriState(prsPart.NotesMasterPart != null);
            prs.HasTitleMaster = TriState(false);
            prs.HasVBProject = TriState(prsPart.VbaProjectPart != null);
            prs.InMergeMode = false;
            prs.LayoutDirection = LayoutDirectionFrom(prsPart);
            prs.NoLineBreakAfter = KinsokuChars(prsPart, start: false);
            prs.NoLineBreakBefore = KinsokuChars(prsPart, start: true);
            prs.Password = "********"; // COM masks password reads with a fixed string
            prs.PasswordEncryptionAlgorithm = string.Empty;
            prs.PasswordEncryptionFileProperties = true;
            prs.PasswordEncryptionKeyLength = 0;
            prs.PasswordEncryptionProvider = string.Empty;
            prs.ReadOnly = 0;
            prs.ReadOnlyRecommended = false;
            prs.RemovePersonalInformation = TriState(prsPart.Presentation?.RemovePersonalInfoOnSave?.Value ?? false);
            prs.Saved = -1;
            prs.SnapToGrid = SnapToGridFrom(prsPart);
            prs.TemplateName = GetAppPropertyValue(document, "Template")
                ?? masterPart?.ThemePart?.Theme?.Name?.Value
                ?? Undefined.String;
            prs.VBASigned = prsPart.VbaProjectPart == null ? 0 : Undefined.Number;
            prs.WritePassword = "********";

            return prs;
        }

        private static SlideMasterPart? FirstMasterPart(PresentationPart prsPart)
        {
            // Master order follows p:sldMasterIdLst, not the part collection order.
            var firstId = prsPart.Presentation?.SlideMasterIdList?.Elements<P.SlideMasterId>().FirstOrDefault();

            if (firstId?.RelationshipId?.Value is string relationshipId
                && prsPart.GetPartById(relationshipId) is SlideMasterPart byId)
            {
                return byId;
            }

            return prsPart.SlideMasterParts.FirstOrDefault();
        }

        private static int LayoutDirectionFrom(PresentationPart prsPart)
        {
            return prsPart.Presentation?.RightToLeft?.Value == true ? 2 : 1; // ppDirectionRightToLeft : ppDirectionLeftToRight
        }

        private static int DefaultLanguageIdFrom(PresentationPart prsPart)
        {
            // The presentation's default editing language lands on the default text
            // style's level-1 run defaults.
            string? language = prsPart.Presentation?.DefaultTextStyle?
                .Descendants<D.DefaultRunProperties>()
                .Select(p => p.Language?.Value)
                .FirstOrDefault(l => !string.IsNullOrEmpty(l));

            if (language == null) return Undefined.Number;

            try
            {
                int lcid = CultureInfo.GetCultureInfo(language).LCID;
                return lcid > 0 && lcid != 4096 ? lcid : Undefined.Number; // 4096 = generated/unknown culture
            }
            catch (CultureNotFoundException)
            {
                return Undefined.Number;
            }
        }

        private static string KinsokuChars(PresentationPart prsPart, bool start)
        {
            var kinsoku = prsPart.Presentation?.GetFirstChild<P.Kinsoku>();
            return (start ? kinsoku?.InvalidStartChars?.Value : kinsoku?.InvalidEndChars?.Value) ?? string.Empty;
        }

        private static string GridDistanceFrom(PresentationPart prsPart)
        {
            long? cx = prsPart.ViewPropertiesPart?.ViewProperties?.GetFirstChild<P.GridSpacing>()?.Cx?.Value;
            return cx.HasValue ? PointsFromEmu(cx) : Undefined.String;
        }

        private static int SnapToGridFrom(PresentationPart prsPart)
        {
            var viewProperties = prsPart.ViewPropertiesPart?.ViewProperties?
                .GetFirstChild<P.SlideViewProperties>()?
                .GetFirstChild<P.CommonSlideViewProperties>();

            return TriState(viewProperties?.SnapToGrid?.Value ?? true);
        }

        // -- Masters ------------------------------------------------------------

        private static PP.SlideMaster BuildMaster(PresentationPart prsPart, SlideMasterPart masterPart,
            string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            var context = OpenXmlPowerPointContext.ForMaster(prsPart, masterPart);

            var master = new PP.SlideMaster
            {
                SlideShowTransition = DefaultSlideShowTransition(),
                SlideMasterLayouts = new List<PP.SlideMasterLayout>(),
                Shapes = BuildShapes(masterPart.SlideMaster?.CommonSlideData?.ShapeTree, context, outputJsonPath, options),
                Name = "SlideMaster", // the COM name of the slide master object
            };

            int index = 0;
            foreach (var layoutPart in LayoutPartsInOrder(masterPart))
            {
                index++;
                var layoutContext = OpenXmlPowerPointContext.ForLayout(prsPart, layoutPart);

                master.SlideMasterLayouts.Add(new PP.SlideMasterLayout
                {
                    SlideShowTransition = DefaultSlideShowTransition(),
                    Shapes = BuildShapes(layoutPart.SlideLayout?.CommonSlideData?.ShapeTree, layoutContext, outputJsonPath, options),
                    Hyperlinks = BuildHyperlinks(layoutPart, layoutPart.SlideLayout),
                    Name = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value ?? Undefined.String,
                    Index = index,
                });
            }

            return master;
        }

        private static List<SlideLayoutPart> LayoutPartsInOrder(SlideMasterPart masterPart)
        {
            // The COM CustomLayouts collection follows p:sldLayoutIdLst order.
            var ordered = new List<SlideLayoutPart>();

            foreach (var layoutId in masterPart.SlideMaster?.SlideLayoutIdList?.Elements<P.SlideLayoutId>()
                ?? Enumerable.Empty<P.SlideLayoutId>())
            {
                if (layoutId.RelationshipId?.Value is string relationshipId
                    && masterPart.GetPartById(relationshipId) is SlideLayoutPart layoutPart)
                {
                    ordered.Add(layoutPart);
                }
            }

            return ordered.Count > 0 ? ordered : masterPart.SlideLayoutParts.ToList();
        }

        private static PP.SlideMaster? BuildNotesMaster(PresentationPart prsPart, string outputJsonPath,
            MsOfficeJsonWriterOptions options)
        {
            var part = prsPart.NotesMasterPart;
            if (part?.NotesMaster?.CommonSlideData == null) return null;

            var context = OpenXmlPowerPointContext.ForNotesMaster(prsPart, part);

            return new PP.SlideMaster
            {
                SlideShowTransition = null,
                SlideMasterLayouts = null,
                Shapes = BuildShapes(part.NotesMaster.CommonSlideData.ShapeTree, context, outputJsonPath, options),
                Name = "NotesMaster",
            };
        }

        private static PP.SlideMaster? BuildHandoutMaster(PresentationPart prsPart, string outputJsonPath,
            MsOfficeJsonWriterOptions options)
        {
            var part = prsPart.HandoutMasterPart;
            if (part?.HandoutMaster?.CommonSlideData == null) return null;

            var context = OpenXmlPowerPointContext.ForHandoutMaster(prsPart, part);

            return new PP.SlideMaster
            {
                // Every transition getter raises on the handout master (per-member nulls).
                SlideShowTransition = new PP.SlideShowTransition(),
                SlideMasterLayouts = null,
                Shapes = BuildShapes(part.HandoutMaster.CommonSlideData.ShapeTree, context, outputJsonPath, options),
                Name = "HandoutMaster",
            };
        }

        // -- Presentation sub-objects ---------------------------------------------

        private static PP.PageSetup BuildPageSetup(PresentationPart prsPart)
        {
            var slideSize = prsPart.Presentation?.SlideSize;
            var notesSize = prsPart.Presentation?.NotesSize;
            long? cx = slideSize?.Cx?.Value;
            long? cy = slideSize?.Cy?.Value;

            int orientation = Undefined.Number;
            if (cx.HasValue && cy.HasValue)
            {
                orientation = cx.Value >= cy.Value ? 1 : 2; // msoOrientationHorizontal = 1, Vertical = 2
            }

            int notesOrientation = Undefined.Number;
            if (notesSize?.Cx?.Value is long notesCx && notesSize.Cy?.Value is long notesCy)
            {
                notesOrientation = notesCx >= notesCy ? 1 : 2;
            }

            return new PP.PageSetup
            {
                FirstSlideNumber = prsPart.Presentation?.FirstSlideNum?.Value ?? 1,
                NotesOrientation = notesOrientation,
                SlideHeight = PointsFromEmu(cy),
                SlideOrientation = orientation,
                SlideSize = PpSlideSizeFrom(slideSize),
                SlideWidth = PointsFromEmu(cx),
            };
        }

        private static int PpSlideSizeFrom(P.SlideSize? slideSize)
        {
            if (slideSize == null) return Undefined.Number;

            var type = slideSize.Type?.Value;
            if (type == null) return 7; // no type attribute is ppSlideSizeCustom

            // PpSlideSizeType values.
            if (type == P.SlideSizeValues.Screen4x3) return 1;
            if (type == P.SlideSizeValues.Letter) return 2;
            if (type == P.SlideSizeValues.A4) return 3;
            if (type == P.SlideSizeValues.Film35mm) return 4;
            if (type == P.SlideSizeValues.Overhead) return 5;
            if (type == P.SlideSizeValues.Banner) return 6;
            if (type == P.SlideSizeValues.Custom) return 7;
            if (type == P.SlideSizeValues.Ledger) return 8;
            if (type == P.SlideSizeValues.A3) return 9;
            if (type == P.SlideSizeValues.B4ISO) return 10;
            if (type == P.SlideSizeValues.B5ISO) return 11;
            if (type == P.SlideSizeValues.B4JIS) return 12;
            if (type == P.SlideSizeValues.B5JIS) return 13;
            if (type == P.SlideSizeValues.HagakiCard) return 14;
            if (type == P.SlideSizeValues.Screen16x9) return 15;
            if (type == P.SlideSizeValues.Screen16x10) return 16;

            return Undefined.Number;
        }

        private static PP.PrintOptions BuildPrintOptions(PresentationPart prsPart)
        {
            var printing = prsPart.PresentationPropertiesPart?.PresentationProperties?
                .GetFirstChild<P.PrintingProperties>();

            // PowerPoint's print-dialog defaults for a freshly opened file; explicit
            // printing properties in the file override the mapped values.
            var printOptions = new PP.PrintOptions
            {
                ActivePrinter = Undefined.String,
                Collate = -1,
                FitToPage = -1,
                FrameSlides = 0,
                HandoutOrder = 2, // ppPrintHandoutHorizontalFirst
                HighQuality = 0,
                NumberOfCopies = 1,
                OutputType = 1, // ppPrintOutputSlides
                PrintColorType = 1, // ppPrintColor
                PrintComments = -1,
                PrintFontsAsGraphics = 0,
                PrintHiddenSlides = -1,
                PrintInBackground = -1,
                RangeType = 1, // ppPrintAll
                SlideShowName = string.Empty,
            };

            if (printing == null) return printOptions;

            if (printing.ScaleToFitPaper?.Value is bool fit) printOptions.FitToPage = TriState(fit);
            if (printing.FrameSlides?.Value is bool frame) printOptions.FrameSlides = TriState(frame);
            if (printing.HiddenSlides?.Value is bool hidden) printOptions.PrintHiddenSlides = TriState(hidden);

            if (printing.PrintWhat?.Value is P.PrintOutputValues output)
            {
                printOptions.OutputType = output switch
                {
                    P.PrintOutputValues v when v == P.PrintOutputValues.Slides => 1,
                    P.PrintOutputValues v when v == P.PrintOutputValues.Handouts1 => 10,
                    P.PrintOutputValues v when v == P.PrintOutputValues.Handouts2 => 2,
                    P.PrintOutputValues v when v == P.PrintOutputValues.Handouts3 => 3,
                    P.PrintOutputValues v when v == P.PrintOutputValues.Handouts4 => 8,
                    P.PrintOutputValues v when v == P.PrintOutputValues.Handouts6 => 4,
                    P.PrintOutputValues v when v == P.PrintOutputValues.Handouts9 => 9,
                    P.PrintOutputValues v when v == P.PrintOutputValues.Notes => 5,
                    P.PrintOutputValues v when v == P.PrintOutputValues.Outline => 6,
                    _ => 1,
                };
            }

            if (printing.ColorMode?.Value is P.PrintColorModeValues mode)
            {
                printOptions.PrintColorType = mode switch
                {
                    P.PrintColorModeValues v when v == P.PrintColorModeValues.Color => 1,
                    P.PrintColorModeValues v when v == P.PrintColorModeValues.Gray => 2,
                    P.PrintColorModeValues v when v == P.PrintColorModeValues.BlackWhite => 3,
                    _ => 1,
                };
            }

            return printOptions;
        }

        private static PP.SlideShowSettings? BuildSlideShowSettings(PresentationPart prsPart)
        {
            var show = prsPart.PresentationPropertiesPart?.PresentationProperties?
                .GetFirstChild<P.ShowProperties>();
            if (show == null) return null;

            var browseMode = show.GetFirstChild<P.BrowseSlideMode>();
            var slideRange = show.GetFirstChild<P.SlideRange>();

            int showType = 1; // ppShowTypeSpeaker
            if (browseMode != null) showType = 2; // ppShowTypeWindow
            if (show.GetFirstChild<P.KioskSlideMode>() != null) showType = 3; // ppShowTypeKiosk

            int rangeType = 1; // ppShowAll
            if (slideRange != null) rangeType = 2; // ppShowSlideRange
            if (show.GetFirstChild<P.CustomShowReference>() != null) rangeType = 3; // ppShowNamedSlideShow

            int showMediaControls = -1;
            var mediaControls = show.Descendants().FirstOrDefault(e => e.LocalName == "showMediaCtrls");
            if (mediaControls?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value is string mediaValue)
            {
                showMediaControls = TriState(mediaValue == "1" || mediaValue == "true");
            }

            return new PP.SlideShowSettings
            {
                PointerColor = BuildPenColor(show.GetFirstChild<P.PenColor>()),
                AdvanceMode = show.UseTimings?.Value ?? true
                    ? 2 // ppSlideShowUseSlideTimings
                    : 1, // ppSlideShowManualAdvance
                EndingSlide = (long?)slideRange?.End?.Value ?? 1,
                LoopUntilStopped = TriState(show.Loop?.Value ?? false),
                RangeType = rangeType,
                ShowMediaControls = showMediaControls,
                ShowPresenterView = -1, // PowerPoint default; not stored in the file format
                ShowScrollbar = TriState(browseMode?.ShowScrollbar?.Value ?? true),
                ShowType = showType,
                ShowWithAnimation = TriState(show.ShowAnimation?.Value ?? true),
                ShowWithNarration = TriState(show.ShowNarration?.Value ?? false),
                StartingSlide = (long?)slideRange?.Start?.Value ?? 1,
            };
        }

        private static PP.ColorFormat? BuildPenColor(P.PenColor? penColor)
        {
            string? hex = penColor?.RgbColorModelHex?.Val?.Value;

            if (hex == null && penColor?.GetFirstChild<D.PresetColor>()?.Val?.InnerText is string preset)
            {
                hex = preset switch
                {
                    "red" => "FF0000",
                    "black" => "000000",
                    "white" => "FFFFFF",
                    "green" => "008000",
                    "blue" => "0000FF",
                    _ => null,
                };
            }

            if (hex == null) return null;

            return new PP.ColorFormat
            {
                Brightness = null,
                Creator = PptCreator,
                ObjectThemeColor = null,
                RGB = OpenXmlPowerPointColors.OleFromHex(hex),
                SchemeColor = 0,
                TintAndShade = null,
                Type = 1,
            };
        }

        private static List<PP.BuiltInDocumentProperty> BuildBuiltInDocumentProperties(PresentationDocument document)
        {
            // PowerPoint's BuiltInDocumentProperties is a fixed 34-entry collection;
            // entries whose COM getter raises for presentations carry a null Value.
            const int typeNumber = 1;
            const int typeDate = 3;
            const int typeString = 4;

            var core = document.PackageProperties;

            static string? Date(DateTime? value)
            {
                return value?.ToLocalTime().ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture);
            }

            string? App(PresentationDocument doc, string elementName)
            {
                return GetAppPropertyValue(doc, elementName);
            }

            var entries = new (string Name, int Type, string? Value)[]
            {
                ("Title", typeString, core.Title ?? string.Empty),
                ("Subject", typeString, core.Subject ?? string.Empty),
                ("Author", typeString, core.Creator ?? string.Empty),
                ("Keywords", typeString, core.Keywords ?? string.Empty),
                ("Comments", typeString, core.Description ?? string.Empty),
                ("Template", typeString, App(document, "Template") ?? string.Empty),
                ("Last author", typeString, core.LastModifiedBy ?? string.Empty),
                ("Revision number", typeString, core.Revision ?? string.Empty),
                ("Application name", typeString, App(document, "Application") ?? string.Empty),
                ("Last print date", typeDate, Date(core.LastPrinted)),
                ("Creation date", typeDate, Date(core.Created)),
                ("Last save time", typeDate, Date(core.Modified)),
                ("Total editing time", typeNumber, App(document, "TotalTime") ?? "0"),
                ("Number of pages", typeNumber, null), // raises for presentations
                ("Number of words", typeNumber, App(document, "Words")),
                ("Number of characters", typeNumber, null),
                ("Security", typeNumber, null),
                ("Category", typeString, core.Category ?? string.Empty),
                ("Format", typeString, App(document, "PresentationFormat") ?? string.Empty),
                ("Manager", typeString, App(document, "Manager") ?? string.Empty),
                ("Company", typeString, App(document, "Company") ?? string.Empty),
                ("Number of bytes", typeNumber, null),
                ("Number of lines", typeNumber, null),
                ("Number of paragraphs", typeNumber, App(document, "Paragraphs")),
                ("Number of slides", typeNumber, App(document, "Slides")),
                // COM recounts notes/hidden slides live: every slide has a notes page.
                ("Number of notes", typeNumber, SlideCount(document).ToString(CultureInfo.InvariantCulture)),
                ("Number of hidden Slides", typeNumber, HiddenSlideCount(document).ToString(CultureInfo.InvariantCulture)),
                ("Number of multimedia clips", typeNumber, App(document, "MMClips")),
                ("Hyperlink base", typeString, App(document, "HyperlinkBase") ?? string.Empty),
                ("Number of characters (with spaces)", typeNumber, null),
                ("Content type", typeString, core.ContentType ?? string.Empty),
                ("Content status", typeString, core.ContentStatus ?? string.Empty),
                ("Language", typeString, core.Language ?? string.Empty),
                ("Document version", typeString, core.Version ?? string.Empty),
            };

            return entries.Select(e => new PP.BuiltInDocumentProperty
            {
                Name = e.Name,
                Creator = PptCreator,
                LinkSource = null,
                LinkToContent = 0,
                Type = e.Type,
                Value = e.Value,
            }).ToList();
        }

        private static int SlideCount(PresentationDocument document)
        {
            return document.PresentationPart?.Presentation?.SlideIdList?.Elements<P.SlideId>().Count() ?? 0;
        }

        private static int HiddenSlideCount(PresentationDocument document)
        {
            var prsPart = document.PresentationPart;
            if (prsPart == null) return 0;

            return SlidePartsInOrder(prsPart).Count(p => p.Slide?.Show?.Value == false);
        }

        private static List<PP.CustomDocumentProperty> BuildCustomDocumentProperties(PresentationDocument document)
        {
            return OpenXmlDocumentProperties.BuildCustom(document.CustomFilePropertiesPart,
                (name, value, type) => new PP.CustomDocumentProperty
                {
                    Name = name,
                    Creator = PptCreator,
                    LinkSource = string.Empty,
                    LinkToContent = 0,
                    Type = type,
                    Value = value switch
                    {
                        // VBA CStr(Boolean) capitalization.
                        "true" or "1" when type == 2 => "True",
                        "false" or "0" when type == 2 => "False",
                        _ => value,
                    },
                });
        }

        private static List<PP.CustomXmlPart>? BuildCustomXmlParts(PresentationDocument document, PresentationPart prsPart)
        {
            var list = new List<PP.CustomXmlPart>();

            // The COM CustomXMLParts collection always starts with the three built-in
            // metadata parts (core, extended and cover-page properties) using fixed IDs.
            var core = document.PackageProperties;

            static string Escape(string? value)
            {
                return new XText(value ?? string.Empty).ToString();
            }

            static string Element(string name, string? value)
            {
                return string.IsNullOrEmpty(value) ? $"<{name}/>" : $"<{name}>{Escape(value)}</{name}>";
            }

            string coreXml = "<cp:coreProperties xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\">"
                + Element("dc:creator", core.Creator)
                + Element("cp:keywords", core.Keywords)
                + Element("dc:description", core.Description)
                + Element("dc:subject", core.Subject)
                + Element("dc:title", core.Title)
                + Element("cp:category", core.Category)
                + Element("cp:contentStatus", core.ContentStatus)
                + "</cp:coreProperties>";

            string extendedXml = "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">"
                + Element("Company", GetAppPropertyValue(document, "Company"))
                + Element("Manager", GetAppPropertyValue(document, "Manager"))
                + "</Properties>";

            const string coverPageXml = "<CoverPageProperties xmlns=\"http://schemas.microsoft.com/office/2006/coverPageProps\">"
                + "<PublishDate/><Abstract/><CompanyAddress/><CompanyPhone/><CompanyFax/><CompanyEmail/>"
                + "</CoverPageProperties>";

            list.Add(new PP.CustomXmlPart
            {
                BuiltIn = true,
                Id = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}",
                NamespaceURI = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                XML = coreXml,
            });
            list.Add(new PP.CustomXmlPart
            {
                BuiltIn = true,
                Id = "{6668398D-A668-4E3E-A5EB-62B293D839F1}",
                NamespaceURI = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
                XML = extendedXml,
            });
            list.Add(new PP.CustomXmlPart
            {
                BuiltIn = true,
                Id = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}",
                NamespaceURI = "http://schemas.microsoft.com/office/2006/coverPageProps",
                XML = coverPageXml,
            });

            // The COM collection lists the file parts in package part order.
            var fileParts = prsPart.CustomXmlParts
                .OrderBy(p => p.Uri?.OriginalString ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var part in fileParts)
            {
                string? xml = null;
                string? namespaceUri = null;

                try
                {
                    using var reader = new StreamReader(part.GetStream());
                    xml = reader.ReadToEnd();

                    var parsed = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
                    namespaceUri = parsed.Root?.Name.NamespaceName;

                    // The COM XML property returns the markup without the XML declaration
                    // and without the XLinq space before self-closing tags ("/>", not " />").
                    xml = parsed.Root?.ToString(SaveOptions.DisableFormatting).Replace(" />", "/>") ?? xml;
                }
                catch
                {
                    // Keep whatever was read.
                }

                list.Add(new PP.CustomXmlPart
                {
                    BuiltIn = false,
                    Id = part.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value?.ToUpperInvariant() ?? Undefined.String,
                    NamespaceURI = namespaceUri ?? Undefined.String,
                    XML = xml ?? Undefined.String,
                });
            }

            return list;
        }

        private static List<PP.Design> BuildDesigns(PresentationPart prsPart)
        {
            var list = new List<PP.Design>();

            var masterIds = prsPart.Presentation?.SlideMasterIdList?.Elements<P.SlideMasterId>().ToList()
                ?? new List<P.SlideMasterId>();

            int index = 0;
            foreach (var masterId in masterIds)
            {
                if (masterId.RelationshipId?.Value is not string relationshipId) continue;
                if (prsPart.GetPartById(relationshipId) is not SlideMasterPart masterPart) continue;

                index++;

                // Legacy title-master semantics: a design reports one when its master
                // carries an explicitly typed "title" layout.
                bool hasTitleLayout = LayoutPartsInOrder(masterPart)
                    .Any(l => l.SlideLayout?.Type?.Value == P.SlideLayoutValues.Title);

                list.Add(new PP.Design
                {
                    Name = masterPart.ThemePart?.Theme?.Name?.Value
                        ?? masterPart.SlideMaster?.CommonSlideData?.Name?.Value
                        ?? Undefined.String,
                    Index = index,
                    HasTitleMaster = TriState(hasTitleLayout),
                    Preserved = TriState(masterPart.SlideMaster?.Preserve?.Value ?? false),
                });
            }

            return list;
        }

        private static List<long>? BuildExtraColors(PresentationPart prsPart)
        {
            var recentColors = prsPart.PresentationPropertiesPart?.PresentationProperties?
                .GetFirstChild<P.ColorMostRecentlyUsed>();
            if (recentColors == null) return null;

            var list = new List<long>();

            foreach (var color in recentColors.Elements<D.RgbColorModelHex>())
            {
                if (OleColorFromHex(color.Val?.Value) is long ole) list.Add(ole);
            }

            return list;
        }

        private static List<PP.Guide>? BuildGuides(PresentationPart prsPart)
        {
            var guideList = prsPart.ViewPropertiesPart?.ViewProperties?
                .GetFirstChild<P.SlideViewProperties>()?
                .GetFirstChild<P.CommonSlideViewProperties>()?
                .GetFirstChild<P.GuideList>();

            if (guideList == null) return null;

            var list = new List<PP.Guide>();

            foreach (var guide in guideList.Elements<P.Guide>())
            {
                list.Add(new PP.Guide
                {
                    // The default guide color PowerPoint reports (a fixed gray).
                    Color = new PP.ColorFormat
                    {
                        Brightness = null,
                        Creator = PptCreator,
                        ObjectThemeColor = null,
                        RGB = 10789796,
                        SchemeColor = 0,
                        TintAndShade = null,
                        Type = 1,
                    },
                    Orientation = guide.Orientation?.Value == P.DirectionValues.Horizontal ? 1 : 2,
                    Position = SingleString((guide.Position?.Value ?? 0) / 8.0),
                });
            }

            return list;
        }

        private static List<PP.PresentationFont>? BuildFonts(PresentationPart prsPart)
        {
            var masterPart = FirstMasterPart(prsPart);
            var fontScheme = masterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;

            var embeddedFontNames = prsPart.Presentation?.GetFirstChild<P.EmbeddedFontList>()?
                .Elements<P.EmbeddedFont>()
                .Select(f => f.Font?.Typeface?.Value)
                .Where(n => n != null)
                .ToList() ?? new List<string?>();

            var names = new List<string>();

            void Add(string? name)
            {
                if (string.IsNullOrEmpty(name) || name![0] == '+') return;
                if (!names.Contains(name)) names.Add(name);
            }

            Add(fontScheme?.MajorFont?.LatinFont?.Typeface?.Value);
            Add(fontScheme?.MinorFont?.LatinFont?.Typeface?.Value);

            IEnumerable<OpenXmlPart> textParts()
            {
                foreach (var slidePart in SlidePartsInOrder(prsPart))
                {
                    yield return slidePart;
                    if (slidePart.NotesSlidePart != null) yield return slidePart.NotesSlidePart;
                }

                foreach (var master in prsPart.SlideMasterParts)
                {
                    yield return master;
                    foreach (var layout in master.SlideLayoutParts) yield return layout;
                }

                if (prsPart.NotesMasterPart != null) yield return prsPart.NotesMasterPart;
                if (prsPart.HandoutMasterPart != null) yield return prsPart.HandoutMasterPart;
            }

            foreach (var part in textParts())
            {
                if (part.RootElement == null) continue;

                // The COM Fonts collection lists fonts used by actual runs (run and
                // paragraph-mark properties), not list-style defaults.
                foreach (var properties in part.RootElement.Descendants()
                    .Where(e => e is D.RunProperties || e is D.EndParagraphRunProperties))
                {
                    foreach (var font in properties.Elements().OfType<D.TextFontType>())
                    {
                        Add(font.Typeface?.Value);
                    }
                }
            }

            return names.Select(name => new PP.PresentationFont
            {
                Name = name,
                Embeddable = -1,
                Embedded = TriState(embeddedFontNames.Contains(name)),
            }).ToList();
        }

        private static List<SlidePart> SlidePartsInOrder(PresentationPart prsPart)
        {
            var list = new List<SlidePart>();

            foreach (var slideId in prsPart.Presentation?.SlideIdList?.Elements<P.SlideId>()
                ?? Enumerable.Empty<P.SlideId>())
            {
                if (slideId.RelationshipId?.Value is string relationshipId
                    && prsPart.GetPartById(relationshipId) is SlidePart slidePart)
                {
                    list.Add(slidePart);
                }
            }

            return list;
        }

        private static List<PP.SectionProperty> BuildSectionProperties(PresentationPart prsPart)
        {
            var list = new List<PP.SectionProperty>();

            var sectionList = prsPart.Presentation?
                .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.SectionList>()
                .FirstOrDefault();

            if (sectionList == null) return list;

            int index = 0;
            foreach (var section in sectionList.Elements<DocumentFormat.OpenXml.Office2010.PowerPoint.Section>())
            {
                index++;
                list.Add(new PP.SectionProperty
                {
                    Index = index,
                    Name = section.Name?.Value ?? string.Empty,
                    SectionID = section.Id?.Value ?? string.Empty,
                    SlidesCount = section.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.SectionSlideIdListEntry>().Count(),
                });
            }

            return list;
        }

        private static Dictionary<uint, int> BuildSectionIndexBySlideId(PresentationPart prsPart)
        {
            var map = new Dictionary<uint, int>();

            var sectionList = prsPart.Presentation?
                .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.SectionList>()
                .FirstOrDefault();

            if (sectionList == null) return map;

            int index = 0;
            foreach (var section in sectionList.Elements<DocumentFormat.OpenXml.Office2010.PowerPoint.Section>())
            {
                index++;
                foreach (var entry in section.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.SectionSlideIdListEntry>())
                {
                    if (entry.Id?.Value is uint slideId) map[slideId] = index;
                }
            }

            return map;
        }

        private static List<PP.Tag>? BuildTags(P.CustomerDataList? customerDataList, OpenXmlPart part)
        {
            if (customerDataList == null) return null;

            var list = new List<PP.Tag>();

            foreach (var child in customerDataList.ChildElements)
            {
                string? relationshipId = child switch
                {
                    P.CustomerData data => data.Id?.Value,
                    P.CustomerDataTags tags => tags.Id?.Value,
                    _ => null,
                };

                if (relationshipId == null) continue;
                if (part.GetPartById(relationshipId) is not UserDefinedTagsPart tagsPart) continue;

                foreach (var tag in tagsPart.TagList?.Elements<P.Tag>() ?? Enumerable.Empty<P.Tag>())
                {
                    list.Add(new PP.Tag
                    {
                        Name = tag.Name?.Value ?? string.Empty,
                        Value = tag.Val?.Value ?? string.Empty,
                    });
                }
            }

            return list.Count == 0 ? null : list;
        }

        // -- Slides ---------------------------------------------------------------

        private static List<PP.Slide> BuildSlides(PresentationPart prsPart, string outputJsonPath,
            MsOfficeJsonWriterOptions options, Action<string>? log)
        {
            var list = new List<PP.Slide>();

            var slideIds = prsPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToList() ?? new List<P.SlideId>();
            int firstSlideNumber = prsPart.Presentation?.FirstSlideNum?.Value ?? 1;
            var sectionIndexBySlideId = BuildSectionIndexBySlideId(prsPart);

            log?.Invoke($"  Writing {slideIds.Count} slides...");

            int index = 0;
            foreach (var slideId in slideIds)
            {
                index++;
                if (index % 10 == 0) log?.Invoke($"  Slide {index}/{slideIds.Count}...");

                string? relationshipId = slideId.RelationshipId?.Value;
                var slidePart = relationshipId == null ? null : (SlidePart?)prsPart.GetPartById(relationshipId);
                if (slidePart == null) continue;

                list.Add(BuildSlide(prsPart, slidePart, slideId, index, firstSlideNumber, sectionIndexBySlideId,
                    outputJsonPath, options));
            }

            return list;
        }

        private static PP.Slide BuildSlide(PresentationPart prsPart, SlidePart slidePart, P.SlideId slideId,
            int slideIndex, int firstSlideNumber, Dictionary<uint, int> sectionIndexBySlideId,
            string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            var layoutPart = slidePart.SlideLayoutPart;
            var masterPart = layoutPart?.SlideMasterPart;
            var context = OpenXmlPowerPointContext.ForSlide(prsPart, slidePart);

            // Default slide names follow the internal slide ID (IDs start at 256).
            uint id = slideId.Id?.Value ?? 0;
            string name = slidePart.Slide?.CommonSlideData?.Name?.Value is string explicitName && explicitName.Length > 0
                ? explicitName
                : "Slide" + (id >= 255 ? id - 255 : id).ToString(CultureInfo.InvariantCulture);

            return new PP.Slide
            {
                SlideShowTransition = BuildSlideShowTransition(slidePart),
                ColorScheme = BuildLegacyColorScheme(context),
                ThemeColorScheme = BuildThemeColorScheme(masterPart?.ThemePart),
                HeadersFooters = BuildHeadersFooters(slidePart),
                NotesPage = BuildNotesPage(prsPart, slidePart, name, slideIndex + firstSlideNumber - 1, outputJsonPath, options),
                Background = BuildBackground(slidePart, context),
                Shapes = BuildShapes(slidePart.Slide?.CommonSlideData?.ShapeTree, context, outputJsonPath, options),
                Tags = BuildTags(slidePart.Slide?.CommonSlideData?.GetFirstChild<P.CustomerDataList>(), slidePart)
                    ?? new List<PP.Tag>(),
                Comments = BuildComments(slidePart),
                Hyperlinks = BuildHyperlinks(slidePart, slidePart.Slide),
                Name = name,
                CustomLayout_Name = layoutPart?.SlideLayout?.CommonSlideData?.Name?.Value ?? Undefined.String,
                Design_Name = masterPart?.ThemePart?.Theme?.Name?.Value ?? Undefined.String,
                SlideID = (int)id,
                SlideIndex = slideIndex,
                SlideNumber = slideIndex + firstSlideNumber - 1,
                sectionIndex = sectionIndexBySlideId.TryGetValue(id, out int sectionIndex)
                    ? sectionIndex
                    : Undefined.Number,
                BackgroundStyle = BackgroundStyleFrom(slidePart),
                DisplayMasterShapes = TriState(slidePart.Slide?.ShowMasterShapes?.Value ?? true),
                FollowMasterBackground = TriState(slidePart.Slide?.CommonSlideData?.Background == null),
                HasNotesPage = TriState(prsPart.NotesMasterPart != null), // COM materializes a notes page for every slide
                Layout = PpSlideLayoutFrom(layoutPart?.SlideLayout?.Type?.Value),
                PrintSteps = 1 + context.Timing.ClickCount,
            };
        }

        private static PP.SlideRange? BuildNotesPage(PresentationPart prsPart, SlidePart slidePart, string name,
            int slideNumber, string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            if (slidePart.NotesSlidePart != null)
            {
                return new PP.SlideRange
                {
                    Shapes = BuildShapes(slidePart.NotesSlidePart.NotesSlide?.CommonSlideData?.ShapeTree,
                        OpenXmlPowerPointContext.ForNotesSlide(prsPart, slidePart.NotesSlidePart), outputJsonPath, options),
                    Name = name,
                };
            }

            // COM materializes a default notes page (from the notes master) for slides
            // that have no notes part yet.
            var notesMasterPart = prsPart.NotesMasterPart;
            var masterTree = notesMasterPart?.NotesMaster?.CommonSlideData?.ShapeTree;
            if (masterTree == null) return null;

            var synthesizedTree = new P.ShapeTree();
            uint shapeId = 1;

            void AddPlaceholder(string type, string namePrefix, bool addTextBody, string? fieldType, string? fieldText)
            {
                var masterShape = masterTree.Elements<P.Shape>()
                    .FirstOrDefault(s => OpenXmlPowerPointContext.PlaceholderOf(s)?.Type?.InnerText == type);
                if (masterShape == null) return;

                var masterPlaceholder = OpenXmlPowerPointContext.PlaceholderOf(masterShape)!;

                shapeId++;

                // Default notes placeholders use body index 1; the slide number keeps
                // the master's own index.
                uint? index = type switch
                {
                    "body" => 1u,
                    "sldNum" => masterPlaceholder.Index?.Value,
                    _ => null,
                };

                var drawingProperties = new P.NonVisualShapeDrawingProperties();
                if (type == "sldImg")
                {
                    drawingProperties.Append(new D.ShapeLocks { NoGrouping = true, NoRotation = true, NoChangeAspect = true });
                }

                var shape = new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties
                        {
                            Id = shapeId,
                            Name = $"{namePrefix} {shapeId - 1}",
                        },
                        drawingProperties,
                        new P.ApplicationNonVisualDrawingProperties(
                            new P.PlaceholderShape
                            {
                                Type = masterPlaceholder.Type?.Value,
                                Index = index,
                            })),
                    new P.ShapeProperties());

                if (addTextBody)
                {
                    var paragraph = new D.Paragraph();
                    if (fieldType != null)
                    {
                        paragraph.Append(new D.Field(new D.Text(fieldText ?? string.Empty))
                        {
                            Id = "{34FF4EF2-2F0A-4E7C-A24F-BD1F46B0E5BF}",
                            Type = fieldType,
                        });
                    }

                    shape.Append(new P.TextBody(new D.BodyProperties(), new D.ListStyle(), paragraph));
                }

                synthesizedTree.Append(shape);
            }

            AddPlaceholder("sldImg", "Slide Image Placeholder", addTextBody: false, fieldType: null, fieldText: null);
            AddPlaceholder("body", "Notes Placeholder", addTextBody: true, fieldType: null, fieldText: null);
            AddPlaceholder("sldNum", "Slide Number Placeholder", addTextBody: true, fieldType: "slidenum",
                fieldText: slideNumber.ToString(CultureInfo.InvariantCulture));

            var context = OpenXmlPowerPointContext.ForSynthesizedNotesSlide(prsPart);

            return new PP.SlideRange
            {
                Shapes = BuildShapes(synthesizedTree, context, outputJsonPath, options),
                Name = name,
            };
        }

        private static int PpSlideLayoutFrom(P.SlideLayoutValues? type)
        {
            // PpSlideLayout values; layouts without a type attribute are custom (the
            // OOXML default).
            if (type == null) return 32; // ppLayoutCustom

            var value = type.Value;

            if (value == P.SlideLayoutValues.Title) return 1;
            if (value == P.SlideLayoutValues.Text) return 2;
            if (value == P.SlideLayoutValues.TwoColumnText) return 3;
            if (value == P.SlideLayoutValues.Table) return 4;
            if (value == P.SlideLayoutValues.TextAndChart) return 5;
            if (value == P.SlideLayoutValues.ChartAndText) return 6;
            if (value == P.SlideLayoutValues.Diagram) return 7; // ppLayoutOrgchart
            if (value == P.SlideLayoutValues.Chart) return 8;
            if (value == P.SlideLayoutValues.TextAndClipArt) return 9;
            if (value == P.SlideLayoutValues.ClipArtAndText) return 10;
            if (value == P.SlideLayoutValues.TitleOnly) return 11;
            if (value == P.SlideLayoutValues.Blank) return 12;
            if (value == P.SlideLayoutValues.TextAndObject) return 13;
            if (value == P.SlideLayoutValues.ObjectAndText) return 14;
            if (value == P.SlideLayoutValues.ObjectOnly) return 15; // ppLayoutLargeObject
            if (value == P.SlideLayoutValues.Object) return 16; // ppLayoutObject
            if (value == P.SlideLayoutValues.TextAndMedia) return 17;
            if (value == P.SlideLayoutValues.MidiaAndText) return 18; // (sic — SDK member name)
            if (value == P.SlideLayoutValues.ObjectOverText) return 19;
            if (value == P.SlideLayoutValues.TextOverObject) return 20;
            if (value == P.SlideLayoutValues.TextAndTwoObjects) return 21;
            if (value == P.SlideLayoutValues.TwoObjectsAndText) return 22;
            if (value == P.SlideLayoutValues.TwoObjectsOverText) return 23;
            if (value == P.SlideLayoutValues.FourObjects) return 24;
            if (value == P.SlideLayoutValues.VerticalText) return 25;
            if (value == P.SlideLayoutValues.ClipArtAndVerticalText) return 26;
            if (value == P.SlideLayoutValues.VerticalTitleAndText) return 27;
            if (value == P.SlideLayoutValues.VerticalTitleAndTextOverChart) return 28;
            if (value == P.SlideLayoutValues.TwoObjects) return 29;
            if (value == P.SlideLayoutValues.ObjectAndTwoObjects) return 30;
            if (value == P.SlideLayoutValues.TwoObjectsAndObject) return 31;
            if (value == P.SlideLayoutValues.Custom) return 32;
            if (value == P.SlideLayoutValues.SectionHeader) return 33;
            if (value == P.SlideLayoutValues.TwoTextAndTwoObjects) return 34; // ppLayoutComparison
            if (value == P.SlideLayoutValues.ObjectText) return 35; // ppLayoutContentWithCaption
            if (value == P.SlideLayoutValues.PictureText) return 36; // ppLayoutPictureWithCaption

            return Undefined.Number;
        }

        private static int BackgroundStyleFrom(SlidePart slidePart)
        {
            // MsoBackgroundStyleIndex; the effective background comes from the slide,
            // then its layout, then the master.
            var background = slidePart.Slide?.CommonSlideData?.Background
                ?? slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Background
                ?? slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.Background;

            if (background == null) return Undefined.Number;

            uint? index = background.GetFirstChild<P.BackgroundStyleReference>()?.Index?.Value;
            if (index >= 1001 && index <= 1012) return (int)(index.Value - 1000); // msoBackgroundStylePreset1..12

            return 0; // msoBackgroundStyleNotAPreset
        }

        private static List<PP.Shape>? BuildBackground(SlidePart slidePart, OpenXmlPowerPointContext context)
        {
            var background = slidePart.Slide?.CommonSlideData?.Background
                ?? slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Background
                ?? slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.Background;

            PP.FillFormat fill = EmptyFill();

            var backgroundProperties = background?.GetFirstChild<P.BackgroundProperties>();
            var styleReference = background?.GetFirstChild<P.BackgroundStyleReference>();

            OpenXmlElement? fillElement = null;
            OpenXmlElement? placeholderColor = null;

            if (backgroundProperties != null)
            {
                fillElement = FirstFillChild(backgroundProperties);
            }
            else if (styleReference?.Index?.Value is uint styleIndex)
            {
                fillElement = ThemeFill(context, styleIndex);
                placeholderColor = OpenXmlPowerPointColors.ColorChoiceOf(styleReference);
            }

            if (fillElement != null)
            {
                fill = FillFromElement(fillElement, context, placeholderColor);
            }

            // The COM Slide.Background is a one-shape range whose only readable members
            // are the fill and a handful of constants.
            var shape = new PP.Shape
            {
                Fill = fill,
                Type = msoAutoShape,
                AutoShapeType = 1,
                BlackWhiteMode = 9, // msoBlackWhiteDontShow
                Creator = PptCreator,
                Tags = null,
            };

            return new List<PP.Shape> { shape };
        }

        private static PP.ColorScheme BuildLegacyColorScheme(OpenXmlPowerPointContext context)
        {
            // The legacy 8-slot scheme PowerPoint synthesizes from the theme: the last
            // two accents map from the hyperlink colors.
            string? Ole(string schemeName)
            {
                string? hex = context.GetThemeColorHex(context.MapSchemeColorName(schemeName));
                return hex == null
                    ? null
                    : OpenXmlPowerPointColors.OleFromHex(hex).ToString(CultureInfo.InvariantCulture);
            }

            return new PP.ColorScheme
            {
                Accent1 = Ole("accent2"),
                Accent2 = Ole("hlink"),
                Accent3 = Ole("folHlink"),
                Background = Ole("bg1"),
                Fill = Ole("accent1"),
                Foreground = Ole("tx1"),
                NotSchemeColor = null,
                SchemeColorMixed = null,
                Shadow = Ole("bg2"),
                Title = Ole("tx2"),
            };
        }

        private static PP.HeadersFooters BuildHeadersFooters(SlidePart slidePart)
        {
            static P.Shape? FindPlaceholder(OpenXmlElement? root, P.PlaceholderValues type)
            {
                return root?.Descendants<P.Shape>()
                    .FirstOrDefault(s => OpenXmlPowerPointContext.PlaceholderOf(s)?.Type?.Value == type);
            }

            PP.HeaderFooter Build(P.PlaceholderValues type)
            {
                var shape = FindPlaceholder(slidePart.Slide, type);

                int format = Undefined.Number;
                if (type == P.PlaceholderValues.DateAndTime)
                {
                    // The date format is inherited from the layout placeholder when the
                    // slide itself carries no date placeholder.
                    var dateShape = shape
                        ?? FindPlaceholder(slidePart.SlideLayoutPart?.SlideLayout, type);
                    format = DateTimeFormatFrom(dateShape?.TextBody?.Descendants<D.Field>().FirstOrDefault()?.Type?.Value);
                }

                // Text/UseFormat (and Format outside date fields) raise on slides.
                return new PP.HeaderFooter
                {
                    Format = format == Undefined.Number ? null : format,
                    Text = null,
                    UseFormat = null,
                    Visible = TriState(shape != null),
                };
            }

            return new PP.HeadersFooters
            {
                DateAndTime = Build(P.PlaceholderValues.DateAndTime),
                Footer = Build(P.PlaceholderValues.Footer),
                Header = null, // slides have no header placeholder (notes/handout pages only)
                SlideNumber = Build(P.PlaceholderValues.SlideNumber),
                DisplayOnTitleSlide = null,
            };
        }

        private static int DateTimeFormatFrom(string? fieldType)
        {
            // PpDateTimeFormat: the datetimeN field types map straight onto the enum.
            if (fieldType == null || !fieldType.StartsWith("datetime", StringComparison.Ordinal)) return Undefined.Number;

            string suffix = fieldType.Substring("datetime".Length);
            if (suffix.Length == 0) return 1; // ppDateTimeMdyy
            if (suffix == "FigureOut") return 14; // what COM reports for the locale-driven field

            return int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out int format) && format >= 1
                ? format
                : Undefined.Number;
        }

        private static PP.SlideShowTransition DefaultSlideShowTransition()
        {
            return new PP.SlideShowTransition
            {
                SoundEffect = null,
                AdvanceOnClick = -1,
                AdvanceOnTime = 0,
                AdvanceTime = "0",
                Duration = "2",
                EntryEffect = 0,
                Hidden = 0,
                LoopSoundUntilNext = 0,
                Speed = 1,
            };
        }

        private static PP.SlideShowTransition BuildSlideShowTransition(SlidePart slidePart)
        {
            // The transition may live inside an mc:AlternateContent choice.
            var transition = slidePart.Slide?.Descendants<P.Transition>().FirstOrDefault();

            bool show = slidePart.Slide?.Show?.Value ?? true;

            var result = DefaultSlideShowTransition();
            result.Hidden = TriState(!show);

            if (transition == null) return result;

            result.AdvanceOnClick = TriState(transition.AdvanceOnClick?.Value ?? true);
            result.AdvanceOnTime = TriState(transition.AdvanceAfterTime?.Value != null);
            result.AdvanceTime = transition.AdvanceAfterTime?.Value is string advance
                ? SingleString(double.Parse(advance, CultureInfo.InvariantCulture) / 1000.0)
                : "0";
            result.EntryEffect = EntryEffectFrom(transition);

            int speed = transition.Speed?.Value switch
            {
                null => 3, // the schema default spd="fast"
                P.TransitionSpeedValues value when value == P.TransitionSpeedValues.Slow => 1,
                P.TransitionSpeedValues value when value == P.TransitionSpeedValues.Medium => 2,
                P.TransitionSpeedValues value when value == P.TransitionSpeedValues.Fast => 3,
                _ => 3,
            };
            result.Speed = speed;

            if (transition.Duration?.Value is string duration)
            {
                result.Duration = SingleString(double.Parse(duration, CultureInfo.InvariantCulture) / 1000.0);
            }
            else
            {
                // Legacy speed values correspond to fixed durations.
                result.Duration = speed switch
                {
                    1 => "1",
                    2 => "0.75",
                    _ => "0.5",
                };
            }

            return result;
        }

        private static int EntryEffectFrom(P.Transition? transition)
        {
            // PpEntryEffect; only the effects with an unambiguous COM value are mapped.
            if (transition == null) return 0; // ppEffectNone

            if (transition.GetFirstChild<P.CutTransition>() is P.CutTransition cut)
            {
                return cut.ThroughBlack?.Value == true ? 258 : 257; // ppEffectCutThroughBlack : ppEffectCut
            }

            if (transition.GetFirstChild<P.FadeTransition>() != null) return 3849; // ppEffectFadeSmoothly
            if (transition.GetFirstChild<P.RandomTransition>() != null) return 513; // ppEffectRandom

            var morph = transition.ChildElements.FirstOrDefault(c => c.LocalName == "morph");
            if (morph != null)
            {
                string? option = morph.GetAttributes().FirstOrDefault(a => a.LocalName == "option").Value;
                return option switch
                {
                    "byWord" => 3955, // ppEffectMorphByWord
                    "byChar" => 3956, // ppEffectMorphByChar
                    _ => 3954, // ppEffectMorphByObject
                };
            }

            if (!transition.HasChildren) return 0;

            return Undefined.Number;
        }

        private static PP.ThemeColorScheme? BuildThemeColorScheme(ThemePart? themePart)
        {
            var colorScheme = themePart?.Theme?.ThemeElements?.ColorScheme;
            if (colorScheme == null) return null;

            static string? Rgb(D.Color2Type? color)
            {
                string? hex = color?.RgbColorModelHex?.Val?.Value;
                if (hex == null && color?.SystemColor?.LastColor?.Value is string lastColor) hex = lastColor;

                long? ole = OleColorFromHex(hex);
                return ole?.ToString(CultureInfo.InvariantCulture);
            }

            return new PP.ThemeColorScheme
            {
                ThemeAccent1 = Rgb(colorScheme.Accent1Color),
                ThemeAccent2 = Rgb(colorScheme.Accent2Color),
                ThemeAccent3 = Rgb(colorScheme.Accent3Color),
                ThemeAccent4 = Rgb(colorScheme.Accent4Color),
                ThemeAccent5 = Rgb(colorScheme.Accent5Color),
                ThemeAccent6 = Rgb(colorScheme.Accent6Color),
                ThemeDark1 = Rgb(colorScheme.Dark1Color),
                ThemeDark2 = Rgb(colorScheme.Dark2Color),
                ThemeFollowedHyperlink = Rgb(colorScheme.FollowedHyperlinkColor),
                ThemeHyperlink = Rgb(colorScheme.Hyperlink),
                ThemeLight1 = Rgb(colorScheme.Light1Color),
                ThemeLight2 = Rgb(colorScheme.Light2Color),
            };
        }

        private static List<PP.Comment>? BuildComments(SlidePart slidePart)
        {
            var commentList = slidePart.SlideCommentsPart?.CommentList;
            if (commentList == null) return null;

            var authors = slidePart.OpenXmlPackage is PresentationDocument doc
                ? doc.PresentationPart?.CommentAuthorsPart?.CommentAuthorList
                : null;

            var list = new List<PP.Comment>();

            foreach (var comment in commentList.Elements<P.Comment>())
            {
                var author = authors?.Elements<P.CommentAuthor>()
                    .FirstOrDefault(a => a.Id?.Value == comment.AuthorId?.Value);

                list.Add(new PP.Comment
                {
                    Author = author?.Name?.Value ?? Undefined.String,
                    AuthorInitials = author?.Initials?.Value ?? Undefined.String,
                    Collapsed = null,
                    DateTime = comment.DateTime?.Value is DateTime dateTime
                        ? dateTime.ToString("MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture)
                        : null,
                    Left = Undefined.String,
                    ProviderID = Undefined.String,
                    Text = comment.Text?.Text ?? string.Empty,
                    Top = Undefined.String,
                    UserID = Undefined.String,
                });
            }

            return list.Count == 0 ? null : list;
        }

        // -- Hyperlinks -------------------------------------------------------------

        private static List<PP.Hyperlink>? BuildHyperlinks(OpenXmlPart part, OpenXmlElement? root)
        {
            if (root == null) return null;

            var list = new List<PP.Hyperlink>();

            foreach (var hyperlink in root.Descendants<D.HyperlinkOnClick>())
            {
                string address = ResolveHyperlinkAddress(part, hyperlink.Id?.Value);

                bool isRange = hyperlink.Parent is D.RunProperties;
                string? runText = isRange ? (hyperlink.Parent?.Parent as D.Run)?.Text?.Text : null;

                list.Add(new PP.Hyperlink
                {
                    Address = address,
                    EmailSubject = EmailSubjectFrom(address),
                    ScreenTip = hyperlink.Tooltip?.Value ?? string.Empty,
                    ShowAndReturn = null,
                    SubAddress = string.Empty,
                    TextToDisplay = isRange ? runText ?? string.Empty : null,
                    Type = isRange ? 0 : 1, // msoHyperlinkRange : msoHyperlinkShape
                    ShapeName = ShapeNameFromAncestors(hyperlink),
                    Text = runText ?? string.Empty,
                });
            }

            return list.Count == 0 ? null : list;
        }

        private static string EmailSubjectFrom(string address)
        {
            if (!address.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) return string.Empty;

            int subjectIndex = address.IndexOf("?subject=", StringComparison.OrdinalIgnoreCase);
            if (subjectIndex < 0) return string.Empty;

            string subject = address.Substring(subjectIndex + "?subject=".Length);
            int ampersandIndex = subject.IndexOf('&');
            return ampersandIndex < 0 ? subject : subject.Substring(0, ampersandIndex);
        }

        private static string ResolveHyperlinkAddress(OpenXmlPart part, string? relationshipId)
        {
            if (string.IsNullOrEmpty(relationshipId)) return string.Empty;

            return part.HyperlinkRelationships
                .FirstOrDefault(r => r.Id == relationshipId)?.Uri?.OriginalString ?? string.Empty;
        }

        private static string ShapeNameFromAncestors(OpenXmlElement element)
        {
            foreach (var ancestor in element.Ancestors())
            {
                string? name = ancestor switch
                {
                    P.Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                    P.Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value,
                    P.GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value,
                    P.ConnectionShape connection => connection.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                    _ => null,
                };

                if (name != null) return LegacyShapeName(name);
            }

            return string.Empty;
        }

        // -- Shapes ---------------------------------------------------------------

        private static List<PP.Shape> BuildShapes(P.ShapeTree? shapeTree, OpenXmlPowerPointContext context,
            string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            var list = new List<PP.Shape>();
            if (shapeTree == null) return list;

            int zOrder = 0;
            AppendShapes(shapeTree.ChildElements, context, list, ref zOrder, outputJsonPath, options);

            return list;
        }

        private static void AppendShapes(IEnumerable<OpenXmlElement> elements, OpenXmlPowerPointContext context,
            List<PP.Shape> list, ref int zOrder, string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            foreach (var element in elements)
            {
                if (element is AlternateContent alternate)
                {
                    var branch = (OpenXmlElement?)alternate.GetFirstChild<AlternateContentChoice>()
                        ?? alternate.GetFirstChild<AlternateContentFallback>();
                    if (branch != null)
                    {
                        AppendShapes(branch.ChildElements, context, list, ref zOrder, outputJsonPath, options);
                    }

                    continue;
                }

                var shape = BuildShapeElement(element, context, ref zOrder, outputJsonPath, options);
                if (shape != null) list.Add(shape);
            }
        }

        private static PP.Shape? BuildShapeElement(OpenXmlElement element, OpenXmlPowerPointContext context,
            ref int zOrder, string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            switch (element)
            {
                case P.Shape shape:
                    zOrder++;
                    return BuildShape(shape, context, zOrder);
                case P.Picture picture:
                    zOrder++;
                    return BuildPicture(picture, context, zOrder, outputJsonPath, options);
                case P.GraphicFrame frame:
                    zOrder++;
                    return BuildGraphicFrame(frame, context, zOrder);
                case P.GroupShape group:
                    zOrder++;
                    return BuildGroupShape(group, context, zOrder, outputJsonPath, options);
                case P.ConnectionShape connection:
                    zOrder++;
                    return BuildConnectionShape(connection, context, zOrder);
                default:
                    return null;
            }
        }

        private static PP.Shape NewShapeEntity(int type, int zOrder)
        {
            return new PP.Shape
            {
                AnimationSettings = null,
                Fill = null,
                Line = null,
                PictureFormat = null,
                PlaceholderFormat = NullPlaceholderFormat(),
                TextFrame = null,
                TextFrame2 = null,
                Image = null,
                ActionSettings = null,
                Tags = new List<PP.Tag>(),
                GroupItems = null,
                Id = Undefined.NumberLong,
                Name = Undefined.String,
                Type = type,
                AutoShapeType = Undefined.Number,
                AlternativeText = Undefined.String,
                Title = Undefined.String,
                Left = Undefined.String,
                Top = Undefined.String,
                Width = Undefined.String,
                Height = Undefined.String,
                Rotation = Undefined.String,
                ZOrderPosition = zOrder,
                Visible = -1,
                BackgroundStyle = 0, // msoBackgroundStyleNotAPreset
                BlackWhiteMode = 1, // msoBlackWhiteAutomatic (spPr@bwMode default "auto")
                Child = 0,
                ConnectionSiteCount = Undefined.NumberLong,
                Connector = 0,
                Creator = PptCreator,
                Decorative = Undefined.Number,
                GraphicStyle = 0,
                HasChart = 0,
                HasInkXML = 0,
                HasSectionZoom = 0,
                HasSmartArt = 0,
                HasTable = 0,
                HasTextFrame = 0,
                HorizontalFlip = 0,
                InkXML = null, // the getter raises unless the shape carries ink
                IsNarration = 0,
                LockAspectRatio = 0,
                ShapeStyle = 0, // msoShapeStyleNotAPreset
                VerticalFlip = 0,
                Vertices = string.Empty,
                MediaType = null, // raises for non-media shapes
            };
        }

        private static PP.PlaceholderFormat NullPlaceholderFormat()
        {
            // Reading PlaceholderFormat members on a non-placeholder raises in COM.
            return new PP.PlaceholderFormat
            {
                ContainedType = null,
                Name = null,
                Position = null,
                Type = null,
            };
        }

        private static string LegacyShapeName(string name)
        {
            int lastSpace = name.LastIndexOf(' ');
            if (lastSpace <= 0) return name;

            string prefix = name.Substring(0, lastSpace);
            string suffix = name.Substring(lastSpace);

            if (!suffix.Skip(1).All(char.IsDigit)) return name;

            return LegacyDefaultShapeNames.TryGetValue(prefix, out string? legacy) ? legacy + suffix : name;
        }

        private static void ApplyNonVisualProperties(PP.Shape entity, P.NonVisualDrawingProperties? properties)
        {
            if (properties == null) return;

            entity.Id = properties.Id?.Value ?? 0;
            entity.Name = LegacyShapeName(properties.Name?.Value ?? string.Empty);
            entity.AlternativeText = properties.Description?.Value ?? string.Empty;
            entity.Title = properties.Title?.Value ?? string.Empty;
            entity.Visible = TriState(!(properties.Hidden?.Value ?? false));
            entity.Decorative = TriState(HasDecorativeExtension(properties));
        }

        private static bool HasDecorativeExtension(P.NonVisualDrawingProperties properties)
        {
            // The adec:decorative marker lives in the cNvPr extension list.
            return properties.Descendants().Any(e => e.LocalName == "decorative"
                && e.GetAttributes().Any(a => a.LocalName == "val" && (a.Value == "1" || a.Value == "true")));
        }

        // Constant/derived members shared by every top-level shape kind.
        private static void ApplyCommonShapeValues(PP.Shape entity, OpenXmlPowerPointContext context,
            P.NonVisualDrawingProperties? nonVisualProperties, bool isPlaceholder)
        {
            uint shapeId = (uint)(entity.Id ?? 0);
            entity.AnimationSettings = BuildAnimationSettings(context, shapeId, isPlaceholder, entity);
            entity.IsNarration = TriState(context.Timing.IsNarrationTarget(shapeId));

            // Reading ActionSettings raises on notes pages and notes/handout masters.
            bool notesContext = context.Kind == OpenXmlPowerPointPageKind.NotesSlide
                || context.Kind == OpenXmlPowerPointPageKind.NotesMaster
                || context.Kind == OpenXmlPowerPointPageKind.HandoutMaster;
            entity.ActionSettings = notesContext
                ? null
                : BuildActionSettings(context, nonVisualProperties, entity.Name ?? string.Empty);
        }

        private static void ApplyTransform(PP.Shape entity, D.Offset? offset, D.Extents? extents, int? rotation, bool? flipH, bool? flipV)
        {
            if (offset != null)
            {
                entity.Left = PointsFromEmu(offset.X?.Value);
                entity.Top = PointsFromEmu(offset.Y?.Value);
            }

            if (extents != null)
            {
                entity.Width = PointsFromEmu(extents.Cx?.Value);
                entity.Height = PointsFromEmu(extents.Cy?.Value);
            }

            entity.Rotation = DegreesFrom60000ths(rotation);
            entity.HorizontalFlip = TriState(flipH ?? false);
            entity.VerticalFlip = TriState(flipV ?? false);
        }

        private static void ApplyShapeTransform(PP.Shape entity, D.Transform2D? xfrm, List<P.Shape> placeholderChain)
        {
            if (xfrm == null)
            {
                xfrm = placeholderChain
                    .Select(s => s.ShapeProperties?.Transform2D)
                    .FirstOrDefault(t => t != null);
            }

            ApplyTransform(entity, xfrm?.Offset, xfrm?.Extents, xfrm?.Rotation?.Value,
                xfrm?.HorizontalFlip?.Value, xfrm?.VerticalFlip?.Value);
        }

        private static PP.Shape BuildShape(P.Shape shape, OpenXmlPowerPointContext context, int zOrder)
        {
            var placeholder = OpenXmlPowerPointContext.PlaceholderOf(shape);
            bool isPlaceholder = placeholder != null;
            bool isFreeform = shape.ShapeProperties?.GetFirstChild<D.CustomGeometry>() != null;
            bool isTextBox = OpenXmlPowerPointTextStyles.IsTextBox(shape);

            int type = isPlaceholder ? msoPlaceholder
                : isTextBox ? msoTextBox
                : isFreeform ? msoFreeform
                : msoAutoShape;

            var entity = NewShapeEntity(type, zOrder);

            ApplyNonVisualProperties(entity, shape.NonVisualShapeProperties?.NonVisualDrawingProperties);

            var chain = placeholder == null ? new List<P.Shape>() : context.GetPlaceholderChain(placeholder);

            ApplyShapeTransform(entity, shape.ShapeProperties?.Transform2D, chain);

            var presetGeometry = shape.ShapeProperties?.GetFirstChild<D.PresetGeometry>();
            string? preset = presetGeometry?.Preset?.InnerText;

            entity.AutoShapeType = preset != null && MsoAutoShapeTypes.TryGetValue(preset, out int autoShapeType)
                ? autoShapeType
                : (preset == null ? msoShapeNotPrimitive : Undefined.Number);

            if (isFreeform)
            {
                entity.AutoShapeType = msoShapeNotPrimitive;
                entity.Vertices = null; // the COM getter returns an array VBA cannot stringify
                entity.ConnectionSiteCount = shape.ShapeProperties?.GetFirstChild<D.CustomGeometry>()?
                    .GetFirstChild<D.ConnectionSiteList>()?.Elements<D.ConnectionSite>().Count() ?? 0;
            }
            else if (preset == null)
            {
                entity.ConnectionSiteCount = 4; // the default rect geometry
            }
            else if (PresetConnectionSites.TryGetValue(preset, out long siteCount))
            {
                entity.ConnectionSiteCount = siteCount;
            }

            entity.LockAspectRatio = TriState(shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties?
                .GetFirstChild<D.ShapeLocks>()?.NoChangeAspect?.Value ?? false);

            entity.Fill = BuildEffectiveFill(shape.ShapeProperties, shape.ShapeStyle, chain, context);
            entity.Line = BuildEffectiveLine(shape.ShapeProperties, shape.ShapeStyle, chain, context);
            entity.ShapeStyle = ShapeStylePresetFrom(shape);

            if (placeholder != null)
            {
                entity.PlaceholderFormat = BuildPlaceholderFormat(placeholder, context,
                    containedType: isTextBox ? msoTextBox : msoAutoShape);
            }

            var textBody = shape.TextBody;
            if (textBody != null)
            {
                // Slide-image placeholders report no text frame even with a txBody.
                entity.HasTextFrame = TriState(placeholder?.Type?.Value != P.PlaceholderValues.SlideImage);
                entity.TextFrame = new PP.TextFrame { TextRange = BuildTextRange(textBody, shape, context) };
                entity.TextFrame2 = new PP.TextFrame2 { TextRange2 = BuildTextRange2(textBody, shape, context) };
            }

            ApplyCommonShapeValues(entity, context, shape.NonVisualShapeProperties?.NonVisualDrawingProperties, isPlaceholder);

            return entity;
        }

        // MsoShapeStyleIndex: a shape reports a gallery preset only while its style
        // references are unmodified. The "Colored Fill - Accent N" row (presets 9-14)
        // is the one combination the gallery stores verbatim.
        private static int ShapeStylePresetFrom(P.Shape shape)
        {
            var style = shape.ShapeStyle;
            if (style == null) return 0; // msoShapeStyleNotAPreset

            var spPr = shape.ShapeProperties;
            bool hasOverrides = FirstFillChild(spPr) != null
                || spPr?.GetFirstChild<D.Outline>() != null
                || spPr?.GetFirstChild<D.EffectList>()?.HasChildren == true;
            if (hasOverrides) return 0;

            if (style.LineReference?.Index?.Value != 2
                || style.FillReference?.Index?.Value != 1
                || style.EffectReference?.Index?.Value != 0)
            {
                return 0;
            }

            static string? Accent(OpenXmlElement? reference)
            {
                string? name = (OpenXmlPowerPointColors.ColorChoiceOf(reference) as D.SchemeColor)?.Val?.InnerText;
                return name != null && name.StartsWith("accent", StringComparison.Ordinal) ? name : null;
            }

            string? accent = Accent(style.FillReference);
            if (accent == null || Accent(style.LineReference) != accent || Accent(style.EffectReference) != accent)
            {
                return 0;
            }

            // The current gallery preset stores exactly one shade(15%) on the line
            // reference; other transforms mean a customized (legacy) style.
            var lineColorTransforms = OpenXmlPowerPointColors.ColorChoiceOf(style.LineReference)?
                .ChildElements.ToList() ?? new List<OpenXmlElement>();
            if (lineColorTransforms.Count != 1 || (lineColorTransforms[0] as D.Shade)?.Val?.Value != 15000)
            {
                return 0;
            }

            if ((OpenXmlPowerPointColors.ColorChoiceOf(style.FontReference) as D.SchemeColor)?.Val?.InnerText != "lt1")
            {
                return 0;
            }

            return int.TryParse(accent.Substring("accent".Length), NumberStyles.Integer, CultureInfo.InvariantCulture,
                out int accentNumber) && accentNumber >= 1 && accentNumber <= 6
                ? 8 + accentNumber // msoShapeStylePreset9..14
                : 0;
        }

        private static PP.PlaceholderFormat BuildPlaceholderFormat(P.PlaceholderShape placeholder,
            OpenXmlPowerPointContext context, int containedType)
        {
            return new PP.PlaceholderFormat
            {
                ContainedType = containedType,
                Name = string.Empty,
                Position = placeholder.Index?.Value ?? 0,
                Type = PpPlaceholderTypeFrom(placeholder.Type?.Value),
            };
        }

        private static int PpPlaceholderTypeFrom(P.PlaceholderValues? type)
        {
            // PpPlaceholderType values; a missing type attribute defaults to obj.
            if (type == null) return 7; // ppPlaceholderObject

            var value = type.Value;

            if (value == P.PlaceholderValues.Title) return 1;
            if (value == P.PlaceholderValues.Body) return 2;
            if (value == P.PlaceholderValues.CenteredTitle) return 3;
            if (value == P.PlaceholderValues.SubTitle) return 4;
            if (value == P.PlaceholderValues.Object) return 7;
            if (value == P.PlaceholderValues.Chart) return 8;
            if (value == P.PlaceholderValues.ClipArt) return 9; // ppPlaceholderBitmap
            if (value == P.PlaceholderValues.Media) return 10; // ppPlaceholderMediaClip
            if (value == P.PlaceholderValues.Diagram) return 11; // ppPlaceholderOrgChart
            if (value == P.PlaceholderValues.Table) return 12;
            if (value == P.PlaceholderValues.SlideNumber) return 13;
            if (value == P.PlaceholderValues.Header) return 14;
            if (value == P.PlaceholderValues.Footer) return 15;
            if (value == P.PlaceholderValues.DateAndTime) return 16;
            if (value == P.PlaceholderValues.Picture) return 18;
            if (value == P.PlaceholderValues.SlideImage) return 1; // COM reports the notes slide image as title

            return Undefined.Number;
        }

        private static PP.Shape BuildPicture(P.Picture picture, OpenXmlPowerPointContext context, int zOrder,
            string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            var applicationProperties = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
            bool isVideo = applicationProperties?.GetFirstChild<D.VideoFromFile>() != null;
            bool isAudio = applicationProperties?.GetFirstChild<D.AudioFromFile>() != null
                || applicationProperties?.GetFirstChild<D.AudioFromCD>() != null;
            bool isSvg = picture.BlipFill?.Blip?.Descendants().Any(e => e.LocalName == "svgBlip") == true;

            var placeholder = applicationProperties?.PlaceholderShape;
            bool isPlaceholder = placeholder != null;

            int type = isPlaceholder ? msoPlaceholder
                : isVideo || isAudio ? msoMedia
                : isSvg ? msoGraphic
                : msoPicture;

            var entity = NewShapeEntity(type, zOrder);

            // PpMediaType (movie = 3, sound = 2); raises for non-media shapes.
            if (isVideo) entity.MediaType = 3;
            else if (isAudio) entity.MediaType = 2;

            ApplyNonVisualProperties(entity, picture.NonVisualPictureProperties?.NonVisualDrawingProperties);

            entity.LockAspectRatio = TriState(picture.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?
                .GetFirstChild<D.PictureLocks>()?.NoChangeAspect?.Value ?? false);

            var chain = placeholder == null ? new List<P.Shape>() : context.GetPlaceholderChain(placeholder);
            ApplyShapeTransform(entity, picture.ShapeProperties?.Transform2D, chain);

            string? preset = picture.ShapeProperties?.GetFirstChild<D.PresetGeometry>()?.Preset?.InnerText;
            entity.AutoShapeType = preset == null
                ? msoShapeNotPrimitive
                : MsoAutoShapeTypes.TryGetValue(preset, out int autoShapeType) ? autoShapeType : 1;
            entity.ConnectionSiteCount = preset != null && PresetConnectionSites.TryGetValue(preset, out long siteCount)
                ? siteCount
                : 4;
            entity.Vertices = null;

            byte[]? imageBytes = ReadImageBytes(picture, context.Part, out string imageExtension);

            if (isSvg)
            {
                // Graphics report their fill from the DrawingML properties when present,
                // else from the SVG content itself (single fill color or the Office
                // MsftOfcThm recolor metadata); multi-color graphics report mixed.
                bool hasFill = FirstFillChild(picture.ShapeProperties) != null
                    || picture.ShapeStyle?.FillReference?.Index?.Value > 0;
                entity.Fill = hasFill
                    ? BuildEffectiveFill(picture.ShapeProperties, picture.ShapeStyle, new List<P.Shape>(), context)
                    : FillFromSvgContent(picture, context) ?? MixedFill();
                entity.Line = MixedLine();
            }
            else
            {
                entity.Fill = BuildEffectiveFill(picture.ShapeProperties, picture.ShapeStyle, new List<P.Shape>(), context);
                entity.Line = BuildEffectiveLine(picture.ShapeProperties, picture.ShapeStyle, new List<P.Shape>(), context);

                entity.PictureFormat = new PP.PictureFormat
                {
                    Brightness = "0.5",
                    ColorType = 1, // msoPictureAutomatic
                    Contrast = "0.5",
                    CropBottom = CropPoints(picture, imageBytes, r => r.Bottom?.Value, vertical: true),
                    CropLeft = CropPoints(picture, imageBytes, r => r.Left?.Value, vertical: false),
                    CropRight = CropPoints(picture, imageBytes, r => r.Right?.Value, vertical: false),
                    CropTop = CropPoints(picture, imageBytes, r => r.Top?.Value, vertical: true),
                    TransparencyColor = -2147483648,
                    // Pictures without an explicit spPr fill report a transparent
                    // background; ones carrying legacy fill markup report 0.
                    TransparentBackground = FirstFillChild(picture.ShapeProperties) == null ? -1 : 0,
                };
            }

            if (placeholder != null)
            {
                entity.PlaceholderFormat = BuildPlaceholderFormat(placeholder, context, containedType: msoPicture);
            }

            entity.Image = BuildImage(imageBytes, imageExtension, entity.Id, outputJsonPath, options);

            ApplyCommonShapeValues(entity, context, picture.NonVisualPictureProperties?.NonVisualDrawingProperties, isPlaceholder);

            return entity;
        }

        private static byte[]? ReadImageBytes(P.Picture picture, OpenXmlPart containerPart, out string extension)
        {
            extension = "bin";

            try
            {
                string? embed = picture.BlipFill?.Blip?.Embed?.Value;
                if (embed == null) return null;

                if (containerPart.GetPartById(embed) is not ImagePart imagePart) return null;

                using var stream = imagePart.GetStream();
                using var memory = new MemoryStream();
                stream.CopyTo(memory);

                extension = ExtensionFromContentType(imagePart.ContentType);
                return memory.ToArray();
            }
            catch
            {
                return null;
            }
        }

        // Crop distances in points: the srcRect percentage applied to the source image
        // size (pixels at the image's native DPI).
        private static string CropPoints(P.Picture picture, byte[]? imageBytes,
            Func<D.SourceRectangle, int?> selector, bool vertical)
        {
            var sourceRectangle = picture.BlipFill?.SourceRectangle;
            int? value = sourceRectangle == null ? null : selector(sourceRectangle);
            if (value == null || value == 0) return "0";

            var dimensions = imageBytes == null ? null : OpenXmlImageDimensions.Parse(imageBytes);
            if (dimensions == null || dimensions.PixelWidth <= 0 || dimensions.PixelHeight <= 0) return Undefined.String;

            // PowerPoint works from the source size in whole EMUs, then converts the
            // cropped distance to points.
            long sizeEmu = (long)((vertical ? dimensions.PixelHeight : dimensions.PixelWidth)
                * 914400.0 / (vertical ? dimensions.DpiY : dimensions.DpiX));
            return SingleString(value.Value / 100000.0 * sizeEmu / 12700.0);
        }

        // Derives the fill a single-color SVG graphic reports: Office stores its
        // recolor as a "MsftOfcThm_<Slot>[_lumMod/_lumOff_<nn>]_Fill" CSS class with
        // the resolved hex; plain single-color graphics report that color as RGB.
        private static PP.FillFormat? FillFromSvgContent(P.Picture picture, OpenXmlPowerPointContext context)
        {
            try
            {
                var svgBlip = picture.BlipFill?.Blip?.Descendants().FirstOrDefault(e => e.LocalName == "svgBlip");
                string? relationshipId = svgBlip?.GetAttributes().FirstOrDefault(a => a.LocalName == "embed").Value;
                if (relationshipId == null) return null;

                if (context.Part.GetPartById(relationshipId) is not ImagePart svgPart) return null;

                using var reader = new StreamReader(svgPart.GetStream());
                string svg = reader.ReadToEnd();

                // Gradient fills (paint server references / gradient stops) make the
                // graphic multi-colored, which COM reports as mixed.
                if (svg.IndexOf("url(#", StringComparison.Ordinal) >= 0
                    || svg.IndexOf("stop-color", StringComparison.Ordinal) >= 0)
                {
                    return null;
                }

                var hexes = System.Text.RegularExpressions.Regex.Matches(svg, "fill\\s*[:=]\\s*[\"']?#([0-9A-Fa-f]{6})")
                    .Cast<System.Text.RegularExpressions.Match>()
                    .Select(m => m.Groups[1].Value.ToUpperInvariant())
                    .Distinct()
                    .ToList();

                if (hexes.Count != 1) return null;

                var fill = NewFillFormat();
                fill.ForeColor = new PP.ColorFormat
                {
                    Brightness = "0",
                    Creator = PptCreator,
                    ObjectThemeColor = 0,
                    RGB = OpenXmlPowerPointColors.OleFromHex(hexes[0]),
                    SchemeColor = 0,
                    TintAndShade = "0",
                    Type = 1,
                };

                var themeClass = System.Text.RegularExpressions.Regex.Match(svg,
                    "MsftOfcThm_([A-Za-z]+[0-9]?)((?:_lum(?:Mod|Off)_[0-9]+)*)_Fill");
                if (themeClass.Success)
                {
                    string? schemeName = themeClass.Groups[1].Value switch
                    {
                        "Accent1" => "accent1",
                        "Accent2" => "accent2",
                        "Accent3" => "accent3",
                        "Accent4" => "accent4",
                        "Accent5" => "accent5",
                        "Accent6" => "accent6",
                        "Text1" => "tx1",
                        "Text2" => "tx2",
                        "Background1" => "bg1",
                        "Background2" => "bg2",
                        "Dark1" => "dk1",
                        "Dark2" => "dk2",
                        "Light1" => "lt1",
                        "Light2" => "lt2",
                        "Hyperlink" => "hlink",
                        "FollowedHyperlink" => "folHlink",
                        _ => null,
                    };

                    if (schemeName != null)
                    {
                        fill.ForeColor.Type = 2;
                        fill.ForeColor.ObjectThemeColor = OpenXmlPowerPointColors.MsoThemeColorIndexFrom(schemeName);
                        fill.ForeColor.SchemeColor = OpenXmlPowerPointColors.LegacySchemeIndexFrom(schemeName);

                        double? lumMod = null;
                        double? lumOff = null;
                        foreach (System.Text.RegularExpressions.Match transform in
                            System.Text.RegularExpressions.Regex.Matches(themeClass.Groups[2].Value, "_lum(Mod|Off)_([0-9]+)"))
                        {
                            double value = double.Parse(transform.Groups[2].Value, CultureInfo.InvariantCulture) / 100.0;
                            if (transform.Groups[1].Value == "Mod") lumMod = value;
                            else lumOff = value;
                        }

                        if (lumMod.HasValue || lumOff.HasValue)
                        {
                            fill.ForeColor.Brightness = lumOff > 0
                                ? SingleString(lumOff.Value)
                                : SingleString((lumMod ?? 1.0) - 1.0);
                        }
                    }
                }

                return fill;
            }
            catch
            {
                return null;
            }
        }

        // The COM writer exports every picture as PNG; convert when possible and keep
        // the original blob when the bytes cannot be decoded (or off Windows).
        private static (byte[] Bytes, string Extension) ConvertToPng(byte[] bytes, string extension)
        {
            if (extension == "png" || extension == "svg") return (bytes, extension);

#if NET
            if (!OperatingSystem.IsWindowsVersionAtLeast(6, 1)) return (bytes, extension);
#endif

            try
            {
                using var source = new MemoryStream(bytes);
                using var image = System.Drawing.Image.FromStream(source);
                using var target = new MemoryStream();

                // Animated GIFs save their first frame.
                image.Save(target, System.Drawing.Imaging.ImageFormat.Png);

                return (target.ToArray(), "png");
            }
            catch
            {
                return (bytes, extension);
            }
        }

        private static ImageBlob? BuildImage(byte[]? bytes, string extension, long? shapeId,
            string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            try
            {
                if (bytes == null) return null;

                (bytes, extension) = ConvertToPng(bytes, extension);

                var blob = new ImageBlob { Extension = extension };

                if (options.BlobOutput == BlobOutput_Enum.Base64)
                {
                    blob.Base64 = Convert.ToBase64String(bytes);
                }
                else
                {
                    string blobFolder = options.GetBlobFolderPath(outputJsonPath);
                    Directory.CreateDirectory(blobFolder);

                    string fileName = $"{shapeId}.{extension}";
                    File.WriteAllBytes(Path.Combine(blobFolder, fileName), bytes);

                    blob.FileName = Path.GetFileName(blobFolder) + "/" + fileName;
                }

                return blob;
            }
            catch
            {
                return null;
            }
        }

        private static PP.Shape BuildGraphicFrame(P.GraphicFrame frame, OpenXmlPowerPointContext context, int zOrder)
        {
            string uri = frame.Graphic?.GraphicData?.Uri?.Value ?? string.Empty;

            int contentType = uri switch
            {
                "http://schemas.openxmlformats.org/drawingml/2006/table" => msoTable,
                "http://schemas.openxmlformats.org/drawingml/2006/chart" => msoChart,
                "http://schemas.openxmlformats.org/drawingml/2006/diagram" => msoSmartArt,
                _ => msoAutoShape,
            };

            var placeholder = frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
            bool isPlaceholder = placeholder != null;

            var entity = NewShapeEntity(isPlaceholder ? msoPlaceholder : contentType, zOrder);

            ApplyNonVisualProperties(entity, frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties);

            var chain = placeholder == null ? new List<P.Shape>() : context.GetPlaceholderChain(placeholder);
            var xfrm = frame.Transform;
            if (xfrm != null)
            {
                ApplyTransform(entity, xfrm.Offset, xfrm.Extents, xfrm.Rotation?.Value, null, null);
            }
            else
            {
                ApplyShapeTransform(entity, null, chain);
            }

            entity.HasTable = TriState(contentType == msoTable);
            entity.HasChart = TriState(contentType == msoChart);
            entity.HasSmartArt = TriState(contentType == msoSmartArt);
            entity.AutoShapeType = msoShapeMixed;
            entity.ConnectionSiteCount = 4;
            entity.Vertices = null;

            if (contentType == msoTable)
            {
                // Tables report a visible fill with a fully mixed fore color, and every
                // line member raises.
                entity.Fill = NewFillFormat();
                entity.Fill.ForeColor = MixedColorFormat();
                entity.Line = NullLineFormat();
            }
            else
            {
                entity.Fill = EmptyFill();
                entity.Line = MixedLine();
            }

            if (placeholder != null)
            {
                entity.PlaceholderFormat = BuildPlaceholderFormat(placeholder, context, containedType: contentType);
            }

            ApplyCommonShapeValues(entity, context, frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties, isPlaceholder);

            return entity;
        }

        private static PP.Shape BuildGroupShape(P.GroupShape group, OpenXmlPowerPointContext context, int zOrder,
            string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            var entity = NewShapeEntity(msoGroup, zOrder);

            ApplyNonVisualProperties(entity, group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties);

            var xfrm = group.GroupShapeProperties?.TransformGroup;
            ApplyTransform(entity, xfrm?.Offset, xfrm?.Extents, xfrm?.Rotation?.Value,
                xfrm?.HorizontalFlip?.Value, xfrm?.VerticalFlip?.Value);

            entity.AutoShapeType = msoShapeMixed;
            entity.ConnectionSiteCount = 0;
            entity.Fill = EmptyFill();
            entity.Line = MixedLine();
            entity.GroupItems = new List<PP.Shape>();

            int childZOrder = 0;
            AppendShapes(group.ChildElements, context, entity.GroupItems, ref childZOrder, outputJsonPath, options);

            foreach (var child in entity.GroupItems)
            {
                child.Child = -1; // msoTrue
            }

            ApplyCommonShapeValues(entity, context, group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties, isPlaceholder: false);

            return entity;
        }

        private static PP.Shape BuildConnectionShape(P.ConnectionShape connection, OpenXmlPowerPointContext context, int zOrder)
        {
            // COM reports connectors as msoAutoShape with Connector = msoTrue.
            var entity = NewShapeEntity(msoAutoShape, zOrder);
            entity.Connector = -1; // msoTrue

            ApplyNonVisualProperties(entity, connection.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties);

            var xfrm = connection.ShapeProperties?.Transform2D;
            ApplyTransform(entity, xfrm?.Offset, xfrm?.Extents, xfrm?.Rotation?.Value,
                xfrm?.HorizontalFlip?.Value, xfrm?.VerticalFlip?.Value);

            string? preset = connection.ShapeProperties?.GetFirstChild<D.PresetGeometry>()?.Preset?.InnerText;

            entity.AutoShapeType = msoShapeMixed;
            entity.ConnectionSiteCount = preset != null && PresetConnectionSites.TryGetValue(preset, out long siteCount)
                ? siteCount
                : 0;
            entity.Vertices = null;

            entity.Fill = MixedFill();
            entity.Line = BuildEffectiveLine(connection.ShapeProperties, connection.ShapeStyle, new List<P.Shape>(), context);

            ApplyCommonShapeValues(entity, context, connection.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
                isPlaceholder: false);

            return entity;
        }

        // -- Animation and action settings ------------------------------------------

        private static PP.AnimationSettings BuildAnimationSettings(OpenXmlPowerPointContext context, uint shapeId,
            bool isPlaceholder, PP.Shape entity)
        {
            bool targeted = context.Timing.TryGetTarget(shapeId, out long order, out bool onClick);

            // Autoshapes with neither a visible fill nor a visible line have no
            // background to animate; pictures and graphic frames always do.
            bool hasBackground = entity.Fill?.Visible == -1
                || entity.Line?.Visible == -1
                || entity.Type == msoPicture
                || entity.Type == msoMedia
                || entity.Type == msoGraphic
                || entity.PictureFormat != null
                || entity.Image != null
                || entity.HasTable == -1
                || entity.HasChart == -1
                || entity.HasSmartArt == -1;

            // PowerPoint materializes animation data for placeholders and animated
            // shapes; everything else reports the not-set sentinel. Master pages never
            // materialize it.
            string advanceTime = !context.IsMasterPage && (targeted || isPlaceholder) ? "0" : AdvanceTimeNotSet;

            string? dimColorHex = context.GetThemeColorHex(context.MapSchemeColorName("folHlink"));

            return new PP.AnimationSettings
            {
                // The default dim color is the legacy scheme's accent 3 slot (the
                // followed-hyperlink theme color).
                DimColor = new PP.ColorFormat
                {
                    Brightness = null,
                    Creator = PptCreator,
                    ObjectThemeColor = null,
                    RGB = dimColorHex == null ? 0 : OpenXmlPowerPointColors.OleFromHex(dimColorHex),
                    SchemeColor = 8, // ppAccent3
                    TintAndShade = null,
                    Type = 2, // msoColorTypeScheme
                },
                PlaySettings = null, // raises for non-media shapes
                SoundEffect = null,
                AdvanceMode = targeted && !onClick ? 2 : 1, // ppAdvanceOnTime : ppAdvanceOnClick
                AdvanceTime = advanceTime,
                AfterEffect = 0, // ppAfterEffectNothing
                Animate = TriState(targeted),
                AnimateBackground = TriState(targeted && hasBackground),
                AnimateTextInReverse = 0,
                AnimationOrder = targeted ? order : null,
                EntryEffect = 257, // ppEffectCut (what COM reports for modern effects too)
                TextLevelEffect = targeted ? 16 : 0, // ppAnimateByAllLevels : ppAnimateLevelNone
                TextUnitEffect = 0,
            };
        }

        private static List<PP.ActionSetting> BuildActionSettings(OpenXmlPowerPointContext context,
            P.NonVisualDrawingProperties? nonVisualProperties, string shapeName)
        {
            PP.ActionSetting Build(int mouseActivation, D.HyperlinkType? hyperlink)
            {
                string address = ResolveHyperlinkAddress(context.Part, hyperlink?.Id?.Value);
                bool hasHyperlink = address.Length > 0;

                return new PP.ActionSetting
                {
                    Hyperlink = !hasHyperlink ? null : new PP.Hyperlink
                    {
                        Address = address,
                        EmailSubject = EmailSubjectFrom(address),
                        ScreenTip = hyperlink?.Tooltip?.Value ?? string.Empty,
                        ShowAndReturn = null,
                        SubAddress = string.Empty,
                        TextToDisplay = null, // raises for shape-level hyperlinks
                        Type = 1, // msoHyperlinkShape
                        ShapeName = shapeName,
                        Text = string.Empty,
                    },
                    SoundEffect = null,
                    MouseActivation = mouseActivation,
                    Action = hasHyperlink ? 7 : 0, // ppActionHyperlink : ppActionNone
                    ActionVerb = null,
                    AnimateAction = TriState(hyperlink?.HighlightClick?.Value ?? false),
                    Run = string.Empty,
                    ShowAndReturn = null,
                    SlideShowName = address, // COM mirrors the hyperlink address here
                };
            }

            return new List<PP.ActionSetting>
            {
                Build(1, nonVisualProperties?.HyperlinkOnClick),
                Build(2, nonVisualProperties?.HyperlinkOnHover),
            };
        }

        // -- Fill and line resolution -------------------------------------------------

        private static readonly Type[] FillElementTypes =
        {
            typeof(D.NoFill), typeof(D.SolidFill), typeof(D.GradientFill),
            typeof(D.BlipFill), typeof(D.PatternFill), typeof(D.GroupFill),
        };

        private static OpenXmlElement? FirstFillChild(OpenXmlElement? parent)
        {
            return parent?.ChildElements.FirstOrDefault(c => FillElementTypes.Any(t => t.IsInstanceOfType(c)));
        }

        private static OpenXmlElement? ThemeFill(OpenXmlPowerPointContext context, uint styleIndex)
        {
            if (styleIndex == 0) return null;

            var formatScheme = context.FormatScheme;
            var styles = (styleIndex >= 1000
                ? (OpenXmlElement?)formatScheme?.BackgroundFillStyleList
                : formatScheme?.FillStyleList)?.ChildElements.ToList();

            int index = (int)(styleIndex >= 1000 ? styleIndex - 1001 : styleIndex - 1);

            return styles != null && index >= 0 && index < styles.Count ? styles[index] : null;
        }

        // The fill values COM reports for shapes with no fill anywhere in the chain.
        private static PP.FillFormat EmptyFill()
        {
            var fill = NewFillFormat();
            fill.BackColor = OpenXmlPowerPointColors.White();
            fill.ForeColor = OpenXmlPowerPointColors.White();
            fill.RotateWithObject = 0;
            fill.Transparency = "1";
            fill.Type = -2; // msoFillMixed
            fill.Visible = 0;
            return fill;
        }

        // The fill values COM reports for shapes without a fill concept (connectors,
        // graphics).
        private static PP.FillFormat MixedFill()
        {
            var fill = NewFillFormat();
            fill.BackColor = OpenXmlPowerPointColors.Mixed();
            fill.ForeColor = OpenXmlPowerPointColors.Mixed();
            fill.RotateWithObject = -2;
            fill.Transparency = "1";
            fill.Type = -2;
            fill.Visible = 0;
            return fill;
        }

        private static PP.FillFormat NewFillFormat()
        {
            return new PP.FillFormat
            {
                BackColor = OpenXmlPowerPointColors.White(),
                ForeColor = OpenXmlPowerPointColors.White(),
                GradientAngle = null, // raises for non-gradient fills
                GradientColorType = -2,
                GradientDegree = null,
                GradientStyle = -2,
                GradientVariant = 0,
                Pattern = -2,
                RotateWithObject = 0,
                TextureAlignment = null,
                TextureHorizontalScale = null,
                TextureName = null,
                TextureOffsetX = null,
                TextureOffsetY = null,
                TextureTile = null,
                TextureType = -2,
                TextureVerticalScale = null,
                Transparency = "0",
                Type = 1,
                Visible = -1,
            };
        }

        private static PP.FillFormat BuildEffectiveFill(OpenXmlElement? shapeProperties, P.ShapeStyle? style,
            List<P.Shape> placeholderChain, OpenXmlPowerPointContext context)
        {
            OpenXmlElement? fillElement = FirstFillChild(shapeProperties);
            OpenXmlElement? placeholderColor = null;
            OpenXmlElement? hiddenFillHost = shapeProperties;

            if (fillElement == null)
            {
                foreach (var inherited in placeholderChain)
                {
                    fillElement = FirstFillChild(inherited.ShapeProperties);
                    if (fillElement != null)
                    {
                        hiddenFillHost = inherited.ShapeProperties;
                        break;
                    }
                }
            }

            if (fillElement == null && style?.FillReference?.Index?.Value is uint fillIndex && fillIndex > 0)
            {
                fillElement = ThemeFill(context, fillIndex);
                placeholderColor = OpenXmlPowerPointColors.ColorChoiceOf(style.FillReference);
            }

            if (fillElement == null && style == null)
            {
                foreach (var inherited in placeholderChain)
                {
                    if (inherited.ShapeStyle?.FillReference?.Index?.Value is uint inheritedIndex && inheritedIndex > 0)
                    {
                        fillElement = ThemeFill(context, inheritedIndex);
                        placeholderColor = OpenXmlPowerPointColors.ColorChoiceOf(inherited.ShapeStyle.FillReference);
                        break;
                    }
                }
            }

            if (fillElement == null) return EmptyFill();

            return FillFromElement(fillElement, context, placeholderColor, hiddenFillHost);
        }

        private static PP.FillFormat FillFromElement(OpenXmlElement fillElement, OpenXmlPowerPointContext context,
            OpenXmlElement? placeholderColor, OpenXmlElement? hiddenFillHost = null)
        {
            var fill = NewFillFormat();

            switch (fillElement)
            {
                case D.NoFill:
                    fill.Visible = 0;

                    // PowerPoint stores the underlying (hidden) fill in a compat extension.
                    var hiddenFill = hiddenFillHost?.Descendants<A14.HiddenFillProperties>().FirstOrDefault();
                    var hiddenSolid = hiddenFill?.GetFirstChild<D.SolidFill>();
                    if (hiddenSolid != null)
                    {
                        fill.ForeColor = OpenXmlPowerPointColors.FromParent(hiddenSolid, context) ?? OpenXmlPowerPointColors.White();
                        fill.Type = 1;
                        fill.Transparency = "0";
                    }
                    else
                    {
                        fill.Type = -2;
                        fill.Transparency = "1";
                    }
                    break;

                case D.SolidFill solid:
                    fill.ForeColor = OpenXmlPowerPointColors.FromParent(solid, context, placeholderColor)
                        ?? OpenXmlPowerPointColors.White();
                    fill.Transparency = OpenXmlPowerPointColors.TransparencyFrom(solid);
                    if (fill.Transparency == "0" && placeholderColor != null)
                    {
                        // Style-matrix fills take their alpha from the reference color.
                        fill.Transparency = OpenXmlPowerPointColors.AlphaTransparencyFrom(placeholderColor);
                    }
                    fill.Type = 1;
                    fill.Visible = -1;
                    break;

                case D.GradientFill gradient:
                    var stops = gradient.GetFirstChild<D.GradientStopList>()?.Elements<D.GradientStop>().ToList();
                    if (stops?.Count > 0)
                    {
                        fill.ForeColor = OpenXmlPowerPointColors.FromParent(stops[0], context, placeholderColor)
                            ?? OpenXmlPowerPointColors.White();
                        fill.BackColor = OpenXmlPowerPointColors.FromParent(stops[stops.Count - 1], context, placeholderColor)
                            ?? OpenXmlPowerPointColors.White();
                    }
                    fill.GradientAngle = DegreesFrom60000ths(gradient.GetFirstChild<D.LinearGradientFill>()?.Angle?.Value);
                    fill.GradientColorType = stops?.Count > 2 ? 4 : 2; // msoGradientMultiColor : msoGradientTwoColors
                    fill.RotateWithObject = TriState(gradient.RotateWithShape?.Value ?? true);
                    fill.Type = 3; // msoFillGradient
                    fill.Visible = -1;
                    break;

                case D.BlipFill blip:
                    fill.RotateWithObject = TriState(blip.RotateWithShape?.Value ?? true);
                    fill.Type = 6; // msoFillPicture
                    fill.Visible = -1;
                    // Picture fills surface through the texture getters.
                    fill.TextureAlignment = -2; // msoTextureAlignmentMixed
                    fill.TextureHorizontalScale = SingleMixed;
                    fill.TextureOffsetX = SingleMixed;
                    fill.TextureOffsetY = SingleMixed;
                    fill.TextureTile = 0;
                    fill.TextureType = 2; // msoTextureUserDefined
                    fill.TextureVerticalScale = SingleMixed;
                    break;

                case D.PatternFill pattern:
                    fill.ForeColor = OpenXmlPowerPointColors.FromParent(pattern.ForegroundColor, context, placeholderColor)
                        ?? OpenXmlPowerPointColors.White();
                    fill.BackColor = OpenXmlPowerPointColors.FromParent(pattern.BackgroundColor, context, placeholderColor)
                        ?? OpenXmlPowerPointColors.White();
                    fill.Type = 2; // msoFillPatterned
                    fill.Visible = -1;
                    break;

                default:
                    fill.Type = -2;
                    fill.Transparency = "1";
                    fill.Visible = 0;
                    break;
            }

            return fill;
        }

        private static PP.ColorFormat MixedColorFormat()
        {
            return new PP.ColorFormat
            {
                Brightness = SingleMixed,
                Creator = PptCreator,
                ObjectThemeColor = -2,
                RGB = -2147483648,
                SchemeColor = -2,
                TintAndShade = SingleMixed,
                Type = -2,
            };
        }

        // The line values COM reports for tables, where every getter raises.
        private static PP.LineFormat NullLineFormat()
        {
            return new PP.LineFormat
            {
                BackColor = new PP.ColorFormat { Creator = PptCreator },
                ForeColor = new PP.ColorFormat { Creator = PptCreator },
                Creator = PptCreator,
            };
        }

        // The line values COM reports when no outline is defined anywhere in the chain.
        private static PP.LineFormat MixedLine()
        {
            return new PP.LineFormat
            {
                BackColor = OpenXmlPowerPointColors.White(),
                ForeColor = OpenXmlPowerPointColors.White(),
                BeginArrowheadLength = -2,
                BeginArrowheadStyle = -2,
                BeginArrowheadWidth = -2,
                Creator = PptCreator,
                DashStyle = -2,
                EndArrowheadLength = -2,
                EndArrowheadStyle = -2,
                EndArrowheadWidth = -2,
                InsetPen = -2,
                Pattern = -2,
                Style = -2,
                Transparency = SingleMixed,
                Visible = 0,
                Weight = SingleMixed,
            };
        }

        private static PP.LineFormat BuildEffectiveLine(OpenXmlElement? shapeProperties, P.ShapeStyle? style,
            List<P.Shape> placeholderChain, OpenXmlPowerPointContext context)
        {
            var candidates = new List<(D.Outline Outline, OpenXmlElement? PlaceholderColor)>();

            if (shapeProperties?.GetFirstChild<D.Outline>() is D.Outline own)
            {
                candidates.Add((own, null));
            }

            foreach (var inherited in placeholderChain)
            {
                if (inherited.ShapeProperties?.GetFirstChild<D.Outline>() is D.Outline chainOutline)
                {
                    candidates.Add((chainOutline, null));
                }
            }

            var lineReference = style?.LineReference
                ?? placeholderChain.Select(s => s.ShapeStyle?.LineReference).FirstOrDefault(r => r != null);
            if (lineReference?.Index?.Value is uint lineIndex && lineIndex > 0)
            {
                var themeLines = context.FormatScheme?.LineStyleList?.Elements<D.Outline>().ToList();
                if (themeLines != null && lineIndex - 1 < themeLines.Count)
                {
                    candidates.Add((themeLines[(int)lineIndex - 1], OpenXmlPowerPointColors.ColorChoiceOf(lineReference)));
                }
            }

            if (candidates.Count == 0) return MixedLine();

            (OpenXmlElement Element, OpenXmlElement? PlaceholderColor)? fill = null;
            foreach (var (outline, placeholderColor) in candidates)
            {
                if (FirstFillChild(outline) is OpenXmlElement fillElement)
                {
                    fill = (fillElement, placeholderColor);
                    break;
                }
            }

            int? widthEmu = candidates.Select(c => c.Outline.Width?.Value).FirstOrDefault(w => w.HasValue);
            var dash = candidates.Select(c => c.Outline.GetFirstChild<D.PresetDash>()).FirstOrDefault(d => d != null);
            var compound = candidates.Select(c => c.Outline.CompoundLineType?.Value)
                .FirstOrDefault(v => v.HasValue);
            var headEnd = candidates.Select(c => c.Outline.GetFirstChild<D.HeadEnd>()).FirstOrDefault(h => h != null);
            var tailEnd = candidates.Select(c => c.Outline.GetFirstChild<D.TailEnd>()).FirstOrDefault(t => t != null);

            var line = new PP.LineFormat
            {
                BackColor = OpenXmlPowerPointColors.White(),
                ForeColor = OpenXmlPowerPointColors.White(),
                BeginArrowheadLength = ArrowheadLengthFrom(headEnd?.Length?.Value),
                BeginArrowheadStyle = ArrowheadStyleFrom(headEnd?.Type?.Value),
                BeginArrowheadWidth = ArrowheadWidthFrom(headEnd?.Width?.Value),
                Creator = PptCreator,
                DashStyle = MsoLineDashStyleFrom(dash?.Val?.Value),
                EndArrowheadLength = ArrowheadLengthFrom(tailEnd?.Length?.Value),
                EndArrowheadStyle = ArrowheadStyleFrom(tailEnd?.Type?.Value),
                EndArrowheadWidth = ArrowheadWidthFrom(tailEnd?.Width?.Value),
                InsetPen = 0,
                Pattern = -2,
                Style = MsoLineStyleFrom(compound),
                Transparency = "0",
                Visible = -1,
                Weight = PointsFromEmu(widthEmu ?? 9525),
            };

            if (fill?.Element is D.NoFill)
            {
                // The underlying (hidden) line lives in a compat extension; without one
                // the invisible line reports the fully mixed values.
                var hiddenLine = shapeProperties?.Descendants<A14.HiddenLineProperties>().FirstOrDefault();
                if (hiddenLine == null) return MixedLine();

                line.Visible = 0;
                if (hiddenLine.Width?.Value is int hiddenWidth) line.Weight = PointsFromEmu(hiddenWidth);
                if (OpenXmlPowerPointColors.FromParent(hiddenLine.GetFirstChild<D.SolidFill>(), context) is PP.ColorFormat hiddenColor)
                {
                    line.ForeColor = hiddenColor;
                }
            }
            else if (fill?.Element is D.SolidFill solidFill)
            {
                line.ForeColor = OpenXmlPowerPointColors.FromParent(solidFill, context, fill.Value.PlaceholderColor)
                    ?? OpenXmlPowerPointColors.White();
                line.Transparency = OpenXmlPowerPointColors.TransparencyFrom(solidFill);
                if (line.Transparency == "0" && fill.Value.PlaceholderColor != null)
                {
                    // Style-matrix lines take their alpha from the reference color.
                    line.Transparency = OpenXmlPowerPointColors.AlphaTransparencyFrom(fill.Value.PlaceholderColor);
                }
            }
            else if (fill?.Element is D.GradientFill gradientFill)
            {
                var firstStop = gradientFill.GetFirstChild<D.GradientStopList>()?.GetFirstChild<D.GradientStop>();
                line.ForeColor = OpenXmlPowerPointColors.FromParent(firstStop, context, fill.Value.PlaceholderColor)
                    ?? OpenXmlPowerPointColors.White();
            }

            return line;
        }

        private static int ArrowheadStyleFrom(D.LineEndValues? type)
        {
            // MsoArrowheadStyle; no element or type none is msoArrowheadNone.
            if (type == null) return 1; // msoArrowheadNone

            var value = type.Value;

            if (value == D.LineEndValues.None) return 1;
            if (value == D.LineEndValues.Triangle) return 2; // msoArrowheadTriangle
            if (value == D.LineEndValues.Arrow) return 3; // msoArrowheadOpen
            if (value == D.LineEndValues.Stealth) return 4; // msoArrowheadStealth
            if (value == D.LineEndValues.Diamond) return 5; // msoArrowheadDiamond
            if (value == D.LineEndValues.Oval) return 6; // msoArrowheadOval

            return Undefined.Number;
        }

        private static int ArrowheadWidthFrom(D.LineEndWidthValues? width)
        {
            // MsoArrowheadWidth; medium is the default.
            if (width == null) return 2; // msoArrowheadWidthMedium

            var value = width.Value;

            if (value == D.LineEndWidthValues.Small) return 1; // msoArrowheadNarrow
            if (value == D.LineEndWidthValues.Medium) return 2;
            if (value == D.LineEndWidthValues.Large) return 3; // msoArrowheadWide

            return Undefined.Number;
        }

        private static int ArrowheadLengthFrom(D.LineEndLengthValues? length)
        {
            // MsoArrowheadLength; medium is the default.
            if (length == null) return 2; // msoArrowheadLengthMedium

            var value = length.Value;

            if (value == D.LineEndLengthValues.Small) return 1; // msoArrowheadShort
            if (value == D.LineEndLengthValues.Medium) return 2;
            if (value == D.LineEndLengthValues.Large) return 3; // msoArrowheadLong

            return Undefined.Number;
        }

        private static int MsoLineDashStyleFrom(D.PresetLineDashValues? dash)
        {
            // MsoLineDashStyle; a missing prstDash is a solid line.
            if (dash == null) return 1; // msoLineSolid

            var value = dash.Value;

            if (value == D.PresetLineDashValues.Solid) return 1;
            if (value == D.PresetLineDashValues.SystemDot) return 2; // msoLineSquareDot
            if (value == D.PresetLineDashValues.Dot) return 3; // msoLineRoundDot
            if (value == D.PresetLineDashValues.Dash) return 4; // msoLineDash
            if (value == D.PresetLineDashValues.SystemDash) return 4;
            if (value == D.PresetLineDashValues.DashDot) return 5; // msoLineDashDot
            if (value == D.PresetLineDashValues.SystemDashDot) return 5;
            if (value == D.PresetLineDashValues.SystemDashDotDot) return 6; // msoLineDashDotDot
            if (value == D.PresetLineDashValues.LargeDash) return 7; // msoLineLongDash
            if (value == D.PresetLineDashValues.LargeDashDot) return 8; // msoLineLongDashDot
            if (value == D.PresetLineDashValues.LargeDashDotDot) return 9; // msoLineLongDashDotDot

            return Undefined.Number;
        }

        private static int MsoLineStyleFrom(D.CompoundLineValues? compound)
        {
            // MsoLineStyle; a missing cmpd attribute is a single line.
            if (compound == null) return 1; // msoLineSingle

            var value = compound.Value;

            if (value == D.CompoundLineValues.Single) return 1;
            if (value == D.CompoundLineValues.Double) return 2; // msoLineThinThin
            if (value == D.CompoundLineValues.ThinThick) return 3; // msoLineThinThick
            if (value == D.CompoundLineValues.ThickThin) return 4; // msoLineThickThin
            if (value == D.CompoundLineValues.Triple) return 5; // msoLineThickBetweenThin

            return Undefined.Number;
        }

        // -- Text -----------------------------------------------------------------

        private static string GetText(P.TextBody textBody)
        {
            var paragraphs = textBody.Elements<D.Paragraph>()
                .Select(paragraph => string.Concat(paragraph.ChildElements.Select(child => child switch
                {
                    D.Run run => run.Text?.Text ?? string.Empty,
                    D.Field field => field.GetFirstChild<D.Text>()?.Text ?? string.Empty,
                    D.Break => "\v",
                    _ => string.Empty,
                })));

            return string.Join("\r", paragraphs);
        }

        private static List<OpenXmlEffectiveRunFormat> BuildEffectiveRunFormats(P.TextBody textBody, P.Shape shape,
            OpenXmlPowerPointContext context)
        {
            var formats = new List<OpenXmlEffectiveRunFormat>();
            var paragraphs = textBody.Elements<D.Paragraph>().ToList();

            for (int index = 0; index < paragraphs.Count; index++)
            {
                var paragraph = paragraphs[index];
                OpenXmlElement? lastRunProperties = null;

                foreach (var child in paragraph.ChildElements)
                {
                    OpenXmlElement? runProperties = child switch
                    {
                        D.Run run => run.RunProperties,
                        D.Field field => field.RunProperties,
                        _ => null,
                    };

                    if (child is D.Run || child is D.Field)
                    {
                        lastRunProperties = runProperties;
                        formats.Add(ResolveRunFormat(runProperties, paragraph, shape, context));
                    }
                }

                // Every paragraph except the last contributes its paragraph mark (the
                // "\r" in the range text), formatted by endParaRPr or the last run.
                // The last paragraph's explicit mark formatting still participates in
                // the Color aggregation.
                if (index < paragraphs.Count - 1)
                {
                    var markProperties = (OpenXmlElement?)paragraph.GetFirstChild<D.EndParagraphRunProperties>()
                        ?? lastRunProperties;
                    formats.Add(ResolveRunFormat(markProperties, paragraph, shape, context));
                }
                else if (formats.Count > 0
                    && paragraph.GetFirstChild<D.EndParagraphRunProperties>() is D.EndParagraphRunProperties lastMark)
                {
                    var markFormat = ResolveRunFormat(lastMark, paragraph, shape, context);
                    markFormat.ColorOnly = true;
                    formats.Add(markFormat);
                }
            }

            if (formats.Count == 0)
            {
                var paragraph = textBody.GetFirstChild<D.Paragraph>();
                if (paragraph != null)
                {
                    formats.Add(ResolveRunFormat(paragraph.GetFirstChild<D.EndParagraphRunProperties>(),
                        paragraph, shape, context));
                }
            }

            return formats;
        }

        private static OpenXmlEffectiveRunFormat ResolveRunFormat(OpenXmlElement? runProperties, D.Paragraph paragraph,
            P.Shape shape, OpenXmlPowerPointContext context)
        {
            var (near, far) = OpenXmlPowerPointTextStyles.BuildChain(runProperties, paragraph, shape, context);
            var chain = near.Concat(far).ToList();

            static D.TextCharacterPropertiesType? Properties(OpenXmlElement element)
            {
                return element as D.TextCharacterPropertiesType;
            }

            var format = new OpenXmlEffectiveRunFormat();

            format.Bold = OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.Bold?.Value) ?? false;
            format.Italic = OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.Italic?.Value) ?? false;
            var ownProperties = runProperties as D.TextCharacterPropertiesType;
            format.BoldExplicitOff = ownProperties?.Bold?.Value == false;
            format.ItalicExplicitOff = ownProperties?.Italic?.Value == false;

            var underline = OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.Underline?.Value);
            format.UnderlineStyle = MsoTextUnderlineFrom(underline);

            var capital = OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.Capital?.Value);
            format.Caps = capital == null || capital == D.TextCapsValues.None
                ? 0 // msoNoCaps
                : capital == D.TextCapsValues.All ? 2 : 1; // msoAllCaps : msoSmallCaps

            var strike = OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.Strike?.Value);
            format.Strike = strike == null || strike == D.TextStrikeValues.NoStrike
                ? 0 // msoNoStrike
                : strike == D.TextStrikeValues.DoubleStrike ? 2 : 1; // msoDoubleStrike : msoSingleStrike

            format.SizePoints = (OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.FontSize?.Value) ?? 1800) / 100.0;

            // Autofit scales the reported font size (PowerPoint rounds to whole points).
            double? fontScale = shape.TextBody?.BodyProperties?.GetFirstChild<D.NormalAutoFit>()?.FontScale?.Value / 100000.0;
            if (fontScale.HasValue && fontScale.Value > 0 && fontScale.Value < 1)
            {
                format.SizePoints = Math.Round(format.SizePoints * fontScale.Value, MidpointRounding.AwayFromZero);
            }

            format.BaselineFraction = (OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.Baseline?.Value) ?? 0) / 100000.0;
            format.SpacingPoints = (OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.Spacing?.Value) ?? 0) / 100.0;
            format.KerningPoints = (OpenXmlPowerPointTextStyles.ResolveValue(chain, e => Properties(e)?.Kerning?.Value) ?? 1200) / 100.0;

            string latin = OpenXmlPowerPointTextStyles.Resolve(chain,
                e => Properties(e)?.GetFirstChild<D.LatinFont>()?.Typeface?.Value) ?? "+mn-lt";
            format.LatinName = ResolveThemeFontName(latin, context);

            format.FarEastName = OpenXmlPowerPointTextStyles.Resolve(chain,
                e => Properties(e)?.GetFirstChild<D.EastAsianFont>()?.Typeface?.Value) ?? "+mn-ea";
            format.ComplexScriptName = OpenXmlPowerPointTextStyles.Resolve(chain,
                e => Properties(e)?.GetFirstChild<D.ComplexScriptFont>()?.Typeface?.Value) ?? "+mn-cs";

            // The style fontRef color slots between the shape-level and inherited sources.
            var colorChoice = OpenXmlPowerPointTextStyles.Resolve(near,
                e => OpenXmlPowerPointColors.ColorChoiceOf(Properties(e)?.GetFirstChild<D.SolidFill>()))
                ?? OpenXmlPowerPointColors.ColorChoiceOf(shape.ShapeStyle?.FontReference)
                ?? OpenXmlPowerPointTextStyles.Resolve(far,
                    e => OpenXmlPowerPointColors.ColorChoiceOf(Properties(e)?.GetFirstChild<D.SolidFill>()));
            format.Color = colorChoice != null
                ? OpenXmlPowerPointColors.FromChoice(colorChoice, context)
                : DefaultTextColor(context);

            return format;
        }

        // Runs without any color in the chain use the tx1 scheme color.
        private static PP.ColorFormat DefaultTextColor(OpenXmlPowerPointContext context)
        {
            string? hex = context.GetThemeColorHex(context.MapSchemeColorName("tx1"));

            return new PP.ColorFormat
            {
                Brightness = "0",
                Creator = PptCreator,
                ObjectThemeColor = 13, // msoThemeColorText1
                RGB = hex == null ? 0 : OpenXmlPowerPointColors.OleFromHex(hex),
                SchemeColor = 2, // ppForeground
                TintAndShade = "0",
                Type = 2,
            };
        }

        private static string ResolveThemeFontName(string typeface, OpenXmlPowerPointContext context)
        {
            if (typeface.Length == 0 || typeface[0] != '+') return typeface;

            var fontScheme = context.FontScheme;
            string? resolved = typeface.StartsWith("+mj", StringComparison.Ordinal)
                ? fontScheme?.MajorFont?.LatinFont?.Typeface?.Value
                : fontScheme?.MinorFont?.LatinFont?.Typeface?.Value;

            return resolved ?? typeface;
        }

        private static PP.TextRange BuildTextRange(P.TextBody textBody, P.Shape shape, OpenXmlPowerPointContext context)
        {
            string text = GetText(textBody);

            return new PP.TextRange
            {
                Font = BuildFont(textBody, shape, context),
                Length = text.Length,
                Start = 1,
                Text = text,
            };
        }

        private static PP.TextRange2 BuildTextRange2(P.TextBody textBody, P.Shape shape, OpenXmlPowerPointContext context)
        {
            string text = GetText(textBody);

            return new PP.TextRange2
            {
                Font = BuildFont2(textBody, shape, context),
                Length = text.Length,
                Start = 1,
                Text = text,
            };
        }

        // Aggregation helpers mirroring the COM range getters: equal values pass
        // through, differing values become the mixed markers.
        private static int AggregateTriState(List<OpenXmlEffectiveRunFormat> formats, Func<OpenXmlEffectiveRunFormat, bool> selector)
        {
            bool first = selector(formats[0]);
            return formats.All(f => selector(f) == first) ? TriState(first) : -2;
        }

        // Like AggregateTriState, but an explicitly-switched-off run mixes with runs
        // that merely inherit the off state (observed COM behavior).
        private static int AggregateOnOff(List<OpenXmlEffectiveRunFormat> formats,
            Func<OpenXmlEffectiveRunFormat, bool> valueSelector, Func<OpenXmlEffectiveRunFormat, bool> explicitOffSelector)
        {
            static int State(bool value, bool explicitOff)
            {
                return value ? 1 : (explicitOff ? 2 : 0);
            }

            int first = State(valueSelector(formats[0]), explicitOffSelector(formats[0]));
            return formats.All(f => State(valueSelector(f), explicitOffSelector(f)) == first)
                ? TriState(first == 1)
                : -2;
        }

        private static int AggregateInt(List<OpenXmlEffectiveRunFormat> formats, Func<OpenXmlEffectiveRunFormat, int> selector)
        {
            int first = selector(formats[0]);
            return formats.All(f => selector(f) == first) ? first : -2;
        }

        private static string AggregateSingle(List<OpenXmlEffectiveRunFormat> formats, Func<OpenXmlEffectiveRunFormat, double> selector)
        {
            // Single-valued range getters report the first run's value even when mixed.
            return SingleString(selector(formats[0]));
        }

        private static string AggregateString(List<OpenXmlEffectiveRunFormat> formats, Func<OpenXmlEffectiveRunFormat, string> selector)
        {
            string first = selector(formats[0]);
            return formats.All(f => selector(f) == first) ? first : string.Empty;
        }

        private static PP.ColorFormat? AggregateColor(List<OpenXmlEffectiveRunFormat> formats)
        {
            static string Key(PP.ColorFormat? color)
            {
                return color == null ? string.Empty : $"{color.Type}|{color.RGB}|{color.ObjectThemeColor}|{color.SchemeColor}|{color.Brightness}|{color.TintAndShade}";
            }

            string first = Key(formats[0].Color);
            if (formats.All(f => Key(f.Color) == first)) return formats[0].Color;

            // Mixed colors report the fully mixed ColorFormat.
            return MixedColorFormat();
        }

        private static PP.Font? BuildFont(P.TextBody textBody, P.Shape shape, OpenXmlPowerPointContext context)
        {
            var formats = BuildEffectiveRunFormats(textBody, shape, context);
            if (formats.Count == 0) return null;

            // Color-only units (the last paragraph mark) join the Color aggregation only.
            var visible = formats.Where(f => !f.ColorOnly).ToList();
            if (visible.Count == 0) visible = formats;

            return new PP.Font
            {
                Color = AggregateColor(formats),
                AutoRotateNumbers = 0,
                BaselineOffset = AggregateSingle(visible, f => f.BaselineFraction),
                Bold = AggregateOnOff(visible, f => f.Bold, f => f.BoldExplicitOff),
                Emboss = 0,
                Italic = AggregateOnOff(visible, f => f.Italic, f => f.ItalicExplicitOff),
                Name = AggregateString(visible, f => f.LatinName),
                NameAscii = AggregateString(visible, f => f.LatinName),
                NameComplexScript = AggregateString(visible, f => f.ComplexScriptName),
                NameFarEast = AggregateString(visible, f => f.FarEastName),
                NameOther = string.Empty,
                Shadow = 0,
                Size = AggregateSingle(visible, f => f.SizePoints),
                Subscript = AggregateTriState(visible, f => f.BaselineFraction < 0),
                Superscript = AggregateTriState(visible, f => f.BaselineFraction > 0),
                Underline = AggregateTriState(visible, f => f.UnderlineStyle != 0),
            };
        }

        private static PP.Font2? BuildFont2(P.TextBody textBody, P.Shape shape, OpenXmlPowerPointContext context)
        {
            var formats = BuildEffectiveRunFormats(textBody, shape, context);
            if (formats.Count == 0) return null;

            // Color-only units (the last paragraph mark) join the Color aggregation only.
            var visible = formats.Where(f => !f.ColorOnly).ToList();
            if (visible.Count == 0) visible = formats;

            return new PP.Font2
            {
                Allcaps = AggregateTriState(visible, f => f.Caps == 2),
                AutoRotateNumbers = 0,
                BaselineOffset = AggregateSingle(visible, f => f.BaselineFraction),
                Bold = AggregateOnOff(visible, f => f.Bold, f => f.BoldExplicitOff),
                Caps = AggregateInt(visible, f => f.Caps),
                DoubleStrikeThrough = AggregateTriState(visible, f => f.Strike == 2),
                Equalize = 0,
                Italic = AggregateOnOff(visible, f => f.Italic, f => f.ItalicExplicitOff),
                Kerning = AggregateSingle(visible, f => f.KerningPoints),
                Name = AggregateString(visible, f => f.LatinName),
                NameAscii = AggregateString(visible, f => f.LatinName),
                NameComplexScript = AggregateString(visible, f => f.ComplexScriptName),
                NameFarEast = AggregateString(visible, f => f.FarEastName),
                NameOther = string.Empty,
                Size = AggregateSingle(visible, f => f.SizePoints),
                Smallcaps = AggregateTriState(visible, f => f.Caps == 1),
                SoftEdgeFormat = 0,
                Spacing = AggregateSingle(visible, f => f.SpacingPoints),
                Strike = AggregateInt(visible, f => f.Strike),
                StrikeThrough = AggregateTriState(visible, f => f.Strike == 1),
                Subscript = AggregateTriState(visible, f => f.BaselineFraction < 0),
                Superscript = AggregateTriState(visible, f => f.BaselineFraction > 0),
                UnderlineStyle = AggregateInt(visible, f => f.UnderlineStyle),
                WordArtformat = -2, // msoTextEffectMixed: not WordArt
            };
        }

        private static int MsoTextUnderlineFrom(D.TextUnderlineValues? underline)
        {
            // MsoTextUnderlineType values.
            if (underline == null) return 0; // msoNoUnderline

            var value = underline.Value;

            if (value == D.TextUnderlineValues.None) return 0;
            if (value == D.TextUnderlineValues.Words) return 1; // msoUnderlineWords
            if (value == D.TextUnderlineValues.Single) return 2; // msoUnderlineSingleLine
            if (value == D.TextUnderlineValues.Double) return 3; // msoUnderlineDoubleLine
            if (value == D.TextUnderlineValues.Heavy) return 4; // msoUnderlineHeavyLine
            if (value == D.TextUnderlineValues.Dotted) return 5; // msoUnderlineDottedLine
            if (value == D.TextUnderlineValues.HeavyDotted) return 6; // msoUnderlineDottedHeavyLine
            if (value == D.TextUnderlineValues.Dash) return 7; // msoUnderlineDashLine
            if (value == D.TextUnderlineValues.DashHeavy) return 8; // msoUnderlineDashHeavyLine
            if (value == D.TextUnderlineValues.DashLong) return 9; // msoUnderlineDashLongLine
            if (value == D.TextUnderlineValues.DashLongHeavy) return 10; // msoUnderlineDashLongHeavyLine
            if (value == D.TextUnderlineValues.DotDash) return 11; // msoUnderlineDotDashLine
            if (value == D.TextUnderlineValues.DotDashHeavy) return 12; // msoUnderlineDotDashHeavyLine
            if (value == D.TextUnderlineValues.DotDotDash) return 13; // msoUnderlineDotDotDashLine
            if (value == D.TextUnderlineValues.DotDotDashHeavy) return 14; // msoUnderlineDotDotDashHeavyLine
            if (value == D.TextUnderlineValues.Wavy) return 15; // msoUnderlineWavyLine
            if (value == D.TextUnderlineValues.WavyHeavy) return 16; // msoUnderlineWavyHeavyLine
            if (value == D.TextUnderlineValues.WavyDouble) return 17; // msoUnderlineWavyDoubleLine

            return Undefined.Number;
        }

        private static string? GetAppPropertyValue(PresentationDocument document, string elementName)
        {
            try
            {
                var part = document.ExtendedFilePropertiesPart;
                if (part == null) return null;

                using var reader = new StreamReader(part.GetStream());
                var root = XDocument.Parse(reader.ReadToEnd()).Root;

                string? value = root?.Elements().FirstOrDefault(e => e.Name.LocalName == elementName)?.Value;
                return string.IsNullOrEmpty(value) ? null : value;
            }
            catch
            {
                return null;
            }
        }
    }
}
